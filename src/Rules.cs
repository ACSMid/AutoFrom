using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;

namespace AutoFrom
{
    public sealed class Sender
    {
        public string Id, Account, Represented;
        public string Alias = "";
        public bool DefaultForAccount;
        public bool IsAlias { get { return !string.IsNullOrEmpty(Alias); } }
        public string FromAddress { get { return IsAlias ? Alias : string.IsNullOrEmpty(Represented) ? Account : Represented; } }
        public string Key { get { return Account + "|" + Represented + "|" + Alias; } }
    }
    public sealed class Rule
    {
        public string Id, SenderId;
        public int Priority;
        public string[] Accounts, Recipients, Domains;
        public bool IncludeSubdomains;
        public bool Matches(string account, string recipient)
        {
            if (Accounts.Length > 0 && !Accounts.Contains(account)) return false;
            if (Recipients.Length == 0 && Domains.Length == 0) return true;
            string domain = recipient.Substring(recipient.LastIndexOf('@') + 1);
            return Recipients.Contains(recipient) || Domains.Any(d => domain == d ||
                (IncludeSubdomains && domain.EndsWith("." + d, StringComparison.Ordinal)));
        }
    }
    public sealed class Config
    {
        public bool Enabled;
        public Sender[] Senders;
        public Rule[] Rules;
        public bool HasRouting { get { return Rules.Length > 0 || Senders.Any(s => s.DefaultForAccount); } }
        public static string Address(string value)
        {
            string s = (value ?? "").Trim().ToLowerInvariant();
            if (s.Length > 254 || s.Count(c => c == '@') != 1 || s.Any(c => c > 127 || char.IsWhiteSpace(c)) ||
                s.IndexOfAny(new[] {'<', '>', ',', ';', '"', '\r', '\n'}) >= 0 || s.StartsWith("@") || s.EndsWith("@"))
                throw new FormatException("Use a plain ASCII email address, not a display name.");
            Domain(s.Substring(s.IndexOf('@') + 1));
            return s;
        }
        private static string Domain(string value)
        {
            string s = (value ?? "").Trim().ToLowerInvariant();
            if (s.Length == 0 || s.Length > 253 || s.Split('.').Any(p => p.Length == 0 || p.Length > 63 ||
                p.StartsWith("-") || p.EndsWith("-") || p.Any(c => !(c >= 'a' && c <= 'z') && !(c >= '0' && c <= '9') && c != '-')))
                throw new FormatException("Invalid domain; use ASCII/punycode without wildcards.");
            return s;
        }
        public static Config Load(string path) { return Parse(File.ReadAllText(path)); }
        public static Config Parse(string json)
        {
            if (json.Length > 131072) throw new FormatException("Configuration exceeds 128 KiB.");
            var root = Obj(new JavaScriptSerializer().DeserializeObject(json), "version", "enabled", "senders", "rules");
            int version = Int(root, "version");
            if (version != 1 && version != 2) throw new FormatException("Unsupported version.");
            var c = new Config { Enabled = Bool(root, "enabled", true) };
            c.Senders = Array(root, "senders").Select(v => {
                var o = Obj(v, "id", "account", "represented", "alias", "defaultForAccount");
                var sender = new Sender { Id = Str(o, "id"), Account = Address(Str(o, "account")),
                    Represented = o.ContainsKey("represented") ? Address(Str(o, "represented")) : "",
                    Alias = o.ContainsKey("alias") ? Address(Str(o, "alias")) : "",
                    DefaultForAccount = Bool(o, "defaultForAccount", false) };
                if (sender.IsAlias && sender.Represented.Length > 0)
                    throw new FormatException("Enter either an Alias or a Shared / delegated mailbox, not both: " + sender.Id);
                if (sender.IsAlias && sender.Alias == sender.Account)
                    throw new FormatException("The alias is the account's primary address. Leave Alias blank to use the account itself.");
                return sender;
            }).ToArray();
            c.Rules = Array(root, "rules").Select(v => {
                var o = Obj(v, "id", "priority", "sender", "accounts", "recipients", "domains", "includeSubdomains");
                var r = new Rule { Id = Str(o, "id"), Priority = Int(o, "priority"), SenderId = Str(o, "sender"),
                    Accounts = Strings(o, "accounts", Address), Recipients = Strings(o, "recipients", Address),
                    Domains = Strings(o, "domains", Domain), IncludeSubdomains = Bool(o, "includeSubdomains", false) };
                if (r.Accounts.Length + r.Recipients.Length + r.Domains.Length == 0)
                    throw new FormatException("Each rule needs at least one account, recipient, or domain condition.");
                return r;
            }).ToArray();
            if (c.Senders.Length > 100 || c.Rules.Length > 500) throw new FormatException("Too many senders or rules.");
            if (c.Senders.Select(s => s.Id).Distinct().Count() != c.Senders.Length ||
                c.Senders.Select(s => s.Key).Distinct().Count() != c.Senders.Length ||
                c.Rules.Select(r => r.Id).Distinct().Count() != c.Rules.Length)
                throw new FormatException("Duplicate sender ID, sender route, or rule ID.");
            var brokenRule = c.Rules.FirstOrDefault(r => !c.Senders.Any(s => s.Id == r.SenderId));
            if (brokenRule != null)
                throw new FormatException("Rule '" + brokenRule.Id + "' uses sender label '" + brokenRule.SenderId +
                    "', which is no longer in Senders. On Rules, choose an existing sender or remove that rule. " +
                    "To use one sender for all mail from its account, remove unnecessary rules and tick Use as default on Senders.");
            if (c.Senders.Where(s => s.DefaultForAccount).GroupBy(s => s.Account).Any(g => g.Count() > 1))
                throw new FormatException("Choose only one default sender for each Outlook account.");
            return c;
        }
        static Dictionary<string, object> Obj(object v, params string[] allowed)
        {
            var d = v as Dictionary<string, object>;
            if (d == null || d.Keys.Any(k => !allowed.Contains(k))) throw new FormatException("Invalid object or unknown configuration field.");
            return d;
        }
        static object[] Array(Dictionary<string, object> d, string k)
        {
            object v; if (!d.TryGetValue(k, out v) || !(v is object[])) throw new FormatException("Missing array: " + k);
            return (object[])v;
        }
        static string Str(Dictionary<string, object> d, string k)
        {
            object v; if (!d.TryGetValue(k, out v) || !(v is string) || string.IsNullOrWhiteSpace((string)v)) throw new FormatException("Missing string: " + k);
            return ((string)v).Trim();
        }
        static int Int(Dictionary<string, object> d, string k)
        {
            object v; if (!d.TryGetValue(k, out v) || !(v is int)) throw new FormatException("Missing integer: " + k);
            return (int)v;
        }
        static bool Bool(Dictionary<string, object> d, string k, bool fallback)
        {
            object v; if (!d.TryGetValue(k, out v)) return fallback;
            if (!(v is bool)) throw new FormatException("Invalid boolean: " + k);
            return (bool)v;
        }
        static string[] Strings(Dictionary<string, object> d, string k, Func<string, string> clean)
        {
            if (!d.ContainsKey(k)) return new string[0];
            return Array(d, k).Select(v => { if (!(v is string)) throw new FormatException("Expected string in " + k); return clean((string)v); }).Distinct().ToArray();
        }
    }
    public sealed class Decision
    {
        public Sender Target;
        public string Error;
    }
    public static class RuleEngine
    {
        public static Decision Decide(Config config, string originalAccount, IList<string> recipients, bool complete)
        {
            if (!config.Enabled || !config.HasRouting) return new Decision();
            var fallback = config.Senders.SingleOrDefault(s => s.DefaultForAccount && s.Account == originalAccount);
            if (recipients.Count == 0 && complete) return new Decision { Target = fallback };
            if (!complete) return new Decision { Error = "Resolve all recipients to individual SMTP addresses. Expand contact groups first." };
            Sender selected = null;
            bool unmatched = false;
            foreach (string address in recipients.Distinct())
            {
                var matches = config.Rules.Where(r => r.Matches(originalAccount, address)).ToArray();
                Sender next = fallback;
                if (matches.Length > 0)
                {
                    int priority = matches.Min(r => r.Priority);
                    var ids = matches.Where(r => r.Priority == priority).Select(r => r.SenderId).Distinct().ToArray();
                    if (ids.Length != 1) return new Decision { Error = "Equal-priority rules select different senders." };
                    next = config.Senders.Single(s => s.Id == ids[0]);
                }
                if (next == null) { unmatched = true; continue; }
                if (selected != null && selected.Id != next.Id)
                    return new Decision { Error = "Recipients require different senders. Split the message or adjust the rules." };
                selected = next;
            }
            if (selected != null && unmatched) return new Decision { Error = "Some recipients have no matching rule. Add an explicit fallback rule or split the message." };
            return new Decision { Target = selected };
        }
    }
}
