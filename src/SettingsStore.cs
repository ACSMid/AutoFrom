using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace AutoFrom
{
    public static class SettingsStore
    {
        public static string DefaultPath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AutoFrom", "rules.json"); } }
        public static Config Empty() { return new Config { Enabled = true, Senders = new Sender[0], Rules = new Rule[0] }; }
        public static string Serialize(Config config)
        {
            var senders = config.Senders.Select(s => {
                var d = new Dictionary<string, object> { {"id", s.Id}, {"account", s.Account} };
                if (!string.IsNullOrWhiteSpace(s.Represented)) d.Add("represented", s.Represented);
                if (s.IsAlias) d.Add("alias", s.Alias);
                if (s.DefaultForAccount) d.Add("defaultForAccount", true);
                return d;
            }).ToArray();
            var rules = config.Rules.Select(r => new {
                id = r.Id, priority = r.Priority, sender = r.SenderId, accounts = r.Accounts,
                recipients = r.Recipients, domains = r.Domains, includeSubdomains = r.IncludeSubdomains
            }).ToArray();
            return new JavaScriptSerializer().Serialize(new { version = 2, enabled = config.Enabled, senders = senders, rules = rules });
        }
        public static Config Validate(Config config) { return Config.Parse(Serialize(config)); }
        public static void Save(string path, Config config)
        {
            string json = Serialize(Validate(config));
            string full = Path.GetFullPath(path);
            Directory.CreateDirectory(Path.GetDirectoryName(full));
            string temporary = full + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    byte[] bytes = new UTF8Encoding(false).GetBytes(json);
                    stream.Write(bytes, 0, bytes.Length);
                    stream.Flush(true);
                }
                if (File.Exists(full)) File.Replace(temporary, full, full + ".bak");
                else File.Move(temporary, full);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
