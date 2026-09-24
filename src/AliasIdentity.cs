using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace AutoFrom
{
    // Documented MAPI address properties. We never modify PR_SENDER_* (the authenticated sender).
    public static class AliasIdentity
    {
        const string Prefix = "http://schemas.microsoft.com/mapi/proptag/";
        public const string ProxyAddresses = Prefix + "0x800F101F";
        public const string AddressType = Prefix + "0x0064001F";
        public const string EmailAddress = Prefix + "0x0065001F";
        public const string SmtpAddress = Prefix + "0x5D02001F";
        public const string EntryId = Prefix + "0x00410102";
        public const string SearchKey = Prefix + "0x003B0102";
        static readonly string[] Properties = { AddressType, EmailAddress, SmtpAddress, EntryId, SearchKey };
        const int NotFound = unchecked((int)0x8004010F);
        public static Dictionary<string, object> Capture(object accessor)
        {
            dynamic pa = accessor;
            var values = new Dictionary<string, object>();
            foreach (string property in Properties)
            {
                try { values[property] = pa.GetProperty(property); }
                catch (COMException ex) { if (ex.ErrorCode != NotFound) throw; }
            }
            return values;
        }
        public static void Restore(object accessor, Dictionary<string, object> values)
        {
            dynamic pa = accessor;
            foreach (string property in Properties)
            {
                object value;
                if (values.TryGetValue(property, out value)) pa.SetProperty(property, value);
                else
                {
                    try { pa.DeleteProperty(property); }
                    catch (COMException ex) { if (ex.ErrorCode != NotFound) throw; }
                }
            }
        }
        public static bool Equal(object left, object right)
        {
            var a = left as byte[]; var b = right as byte[];
            return a != null && b != null ? a.SequenceEqual(b) : object.Equals(left, right);
        }
        public static bool Matches(Dictionary<string, object> actual, Dictionary<string, object> expected)
        {
            return actual != null && actual.Count == expected.Count && expected.All(p => actual.ContainsKey(p.Key) && Equal(actual[p.Key], p.Value));
        }
        public static bool MatchesAlias(Dictionary<string, object> actual, Dictionary<string, object> expected)
        {
            if (Matches(actual, expected)) return true;
            // Observed Outlook Save normalization: both legacy routing strings become
            // empty, while the literal SMTP and complete one-off identity stay intact.
            // Accept only this exact pattern, not missing/partial/conflicting fields.
            if (actual == null || expected == null || actual.Count != expected.Count) return false;
            object type, address;
            if (!actual.TryGetValue(AddressType, out type) || !object.Equals(type, "") ||
                !actual.TryGetValue(EmailAddress, out address) || !object.Equals(address, "")) return false;
            return new[] { SmtpAddress, EntryId, SearchKey }.All(key =>
                actual.ContainsKey(key) && expected.ContainsKey(key) && Equal(actual[key], expected[key]));
        }
        public static string Differences(Dictionary<string, object> actual, Dictionary<string, object> expected)
        {
            var lines = new List<string>();
            foreach (var pair in expected)
            {
                object value;
                bool present = actual != null && actual.TryGetValue(pair.Key, out value);
                value = null;
                if (present) value = actual[pair.Key];
                if (!present || !Equal(value, pair.Value))
                    lines.Add(pair.Key.Substring(Prefix.Length) + ": expected " + Describe(pair.Value) + "; read back " + (present ? Describe(value) : "<missing>"));
            }
            return string.Join("\n", lines);
        }
        static string Describe(object value)
        {
            var bytes = value as byte[];
            if (bytes != null) return "binary[" + bytes.Length + "] " + BitConverter.ToString(bytes.Take(160).ToArray()) + (bytes.Length > 160 ? "..." : "");
            return value == null ? "<null>" : "[" + value.GetType().Name + "] \"" + Convert.ToString(value) + "\"";
        }
        public static void ValidateOwnership(object accountObject, string accountAddress, string alias)
        {
            dynamic account = accountObject;
            if ((int)account.AccountType != 0) throw new InvalidOperationException("Alias sending requires the alias owner's Exchange Online account.");
            object proxies;
            try
            {
                dynamic entry = account.CurrentUser.AddressEntry;
                proxies = entry.PropertyAccessor.GetProperty(ProxyAddresses);
            }
            catch (Exception ex) { throw new InvalidOperationException("Cannot read this account's email aliases. Refresh Outlook's address book and try again. " + ex.Message); }
            var values = proxies as Array;
            if (values == null) throw new InvalidOperationException("Outlook did not return an alias list for this account. Refresh its address book and try again.");
            var smtpAddresses = new HashSet<string>(values.Cast<object>().OfType<string>()
                .Where(p => p.StartsWith("smtp:", StringComparison.OrdinalIgnoreCase)).Select(p => p.Substring(5)), StringComparer.OrdinalIgnoreCase);
            // Account.SmtpAddress can itself be an alias or an older profile address. The
            // account's CurrentUser is authoritative; require BOTH configured addresses in
            // that same object's proxy list rather than insisting on its primary SMTP.
            if (!smtpAddresses.Contains(accountAddress))
                throw new InvalidOperationException("The Outlook account address '" + accountAddress +
                    "' is not in the address list for this account's current mailbox user. Choose the owning mailbox account or refresh Outlook's address book.");
            if (!smtpAddresses.Contains(alias))
                throw new InvalidOperationException("Alias '" + alias + "' is not listed on the mailbox for Outlook account '" + accountAddress +
                    "'. Choose its owning account or refresh the address book.");
        }
        public static Dictionary<string, object> For(string alias)
        {
            alias = Config.Address(alias);
            return new Dictionary<string, object> {
                {AddressType, "SMTP"}, {EmailAddress, alias}, {SmtpAddress, alias},
                {EntryId, OneOff(alias)}, {SearchKey, Encoding.ASCII.GetBytes("SMTP:" + alias.ToUpperInvariant() + "\0")}
            };
        }
        public static byte[] OneOff(string address)
        {
            address = Config.Address(address);
            // MS-OXCDATA 2.2.5.1. Provider UID is a byte sequence, not a .NET Guid.
            // 01 90 = MIME, UTF-16, no directory lookup. No-lookup prevents resolving the
            // alias back to the primary address. Ownership is checked separately above.
            byte[] header = { 0,0,0,0, 0x81,0x2B,0x1F,0xA4,0xBE,0xA3,0x10,0x19,
                0x9D,0x6E,0,0xDD,1,0x0F,0x54,2, 0,0, 1,0x90 };
            using (var stream = new MemoryStream())
            {
                stream.Write(header, 0, header.Length);
                foreach (string field in new[] { address, "SMTP", address })
                {
                    byte[] bytes = Encoding.Unicode.GetBytes(field + "\0");
                    stream.Write(bytes, 0, bytes.Length);
                }
                return stream.ToArray();
            }
        }
    }
}
