using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Linq;
using System.Windows.Forms;
using AutoFrom;

class Tests
{
    static int count;
    static void Check(bool condition, string name) { if (!condition) throw new Exception("FAIL: " + name); count++; Console.WriteLine("PASS: " + name); }
    static void Throws(Action action, string name) { bool threw = false; try { action(); } catch { threw = true; } Check(threw, name); }
    static Config Example;
    static Decision Decide(params string[] recipients) { return RuleEngine.Decide(Example, "alex@contoso.com", recipients, true); }
    [STAThread]
    public static int Main(string[] args)
    {
        try
        {
            Check(Config.Load(args[0]).Rules.Length == 0, "initial configuration is inert");
            Example = Config.Load(args[1]);
            Check(Decide("buyer@partner.example").Target.Id == "second-account", "exact recipient overrides account fallback");
            Check(Decide("one@customer.example").Target.Id == "sales", "domain routing");
            Check(Decide("one@other.example").Target.Id == "personal", "original account fallback");
            Check(Decide("one@evilcustomer.example").Target.Id == "personal", "domain boundary");
            Check(Decide("one@sub.customer.example").Target.Id == "personal", "subdomains require opt-in");
            Example.Rules[1].IncludeSubdomains = true;
            Check(Decide("one@sub.customer.example").Target.Id == "sales", "subdomain opt-in");
            Check(Decide("one@customer.example", "buyer@partner.example").Error != null, "cross-recipient conflict despite priorities");
            Check(Decide("one@customer.example", "two@customer.example").Target.Id == "sales", "same sender across recipients");
            Check(Decide("one@customer.example", "one@customer.example").Target.Id == "sales", "duplicate recipient");
            Check(Decide().Target == null, "empty draft has no mapping");
            Check(RuleEngine.Decide(Example, "alex@contoso.com", new string[0], false).Error != null, "unresolved/group recipient blocks");
            Check(RuleEngine.Decide(Example, "alex@fabrikam.com", new[] {"one@customer.example"}, true).Target == null, "account condition limits rule scope");
            Example.Rules[2].Priority = 20;
            Check(Decide("one@customer.example").Error != null, "equal-priority conflict");
            Example = Config.Load(args[1]);
            Example.Rules = new[] { Example.Rules[1] };
            Check(Decide("one@customer.example", "one@other.example").Error != null, "partially covered message blocks");
            Check(Decide("one@other.example").Target == null, "wholly unmatched message unchanged");
            Example.Enabled = false;
            Check(Decide("one@customer.example").Target == null, "disabled configuration");
            Example.Enabled = true;
            var fake = new Fake(); var state = new DraftState();
            Controller.Evaluate(Example, fake, state, false);
            Check(fake.Current.Represented == "sales@contoso.com", "compose changes real sender port");
            fake.Current.Account = "manual@other.example";
            Controller.Evaluate(Example, fake, state, true);
            Check(fake.Current.Account == "alex@contoso.com" && fake.Resolved, "send re-evaluates and overrides manual account change");
            fake.Addresses = new[] {"unmatched@elsewhere.example"};
            Controller.Evaluate(Example, fake, state, false);
            Check(fake.Current.Represented == "" && !state.Managed, "removing match restores original sender");
            fake.Addresses = new[] {"one@customer.example"}; fake.FailSet = true;
            Throws(() => Controller.Evaluate(Example, fake, state, true), "write error propagates to send cancellation");
            fake.FailSet = false; fake.Addresses = new[] {"one@other.example"};
            Controller.Evaluate(Example, fake, state, true);
            Check(fake.Current.Represented == "", "partial write restored when match disappears");
            fake.Addresses = new[] {"one@customer.example"};
            Controller.Evaluate(Example, fake, state, false);
            fake.Addresses = new[] {"one@other.example"}; fake.FailRestore = true;
            Throws(() => Controller.Evaluate(Example, fake, state, true), "failed restoration blocks send");
            Throws(() => Config.Parse("{}"), "missing version rejected");
            Throws(() => Config.Parse("{\"version\":1,\"senders\":[],\"rules\":[],\"typo\":1}"), "unknown fields rejected");
            Throws(() => Config.Parse("{\"version\":1,\"senders\":[],\"rules\":[{\"id\":\"r\",\"priority\":0,\"sender\":\"missing\",\"domains\":[\"a.test\"]}]}"), "unknown sender rejected");
            Throws(() => Config.Address("Person <one@example.com>"), "display names rejected");
            Check(Config.Address(" ONE@Example.COM ") == "one@example.com", "case normalization");
            var session = new MockSession(); var mail = new MockMail { SendUsingAccount = session.Accounts.Items[0] };
            var port = new OutlookPort(mail, session);
            port.Set(new Sender { Id = "second", Account = "alex@fabrikam.com", Represented = "" });
            Check(mail.SendUsingAccount.SmtpAddress == "alex@fabrikam.com", "Outlook adapter sets SendUsingAccount");
            port.Set(new Sender { Id = "shared", Account = "alex@contoso.com", Represented = "sales@contoso.com" });
            Check(mail.SentOnBehalfOfName == "sales@contoso.com" && mail.SendUsingAccount.SmtpAddress == "alex@contoso.com", "adapter sets delegated sender and transport together");
            var saved = port.Capture();
            port.Set(new Sender { Id = "second", Account = "alex@fabrikam.com", Represented = "" });
            Check(mail.SentOnBehalfOfName == "", "account-only route clears delegated sender");
            port.Restore(saved);
            Check(mail.SentOnBehalfOfName == "sales@contoso.com", "adapter restores delegated snapshot");
            Throws(() => port.Set(new Sender { Account = "missing@example.com", Represented = "" }), "adapter rejects missing profile account");
            Throws(() => port.Set(new Sender { Account = "alex@fabrikam.com", Represented = "sales@contoso.com" }), "adapter rejects delegation through non-Exchange account");
            session.PrimaryAddress = "another@contoso.com";
            Throws(() => port.Set(new Sender { Account = "alex@contoso.com", Represented = "alias@contoso.com" }), "adapter rejects non-primary delegated alias");
            session.PrimaryAddress = "sales@contoso.com";
            bool complete; port.Recipients(true, out complete);
            Check(complete && mail.Recipients.ResolveCalled, "adapter resolves To Cc Bcc collection at send");
            mail.Recipients.Items[0].AddressEntry.AddressEntryUserType = 1;
            port.Recipients(false, out complete);
            Check(!complete, "adapter rejects distribution list rather than skipping it");
            mail.SendUsingAccount = null;
            Throws(() => port.Capture(), "null sending account never guessed");
            var configured = Config.Load(args[1]);
            var roundtrip = SettingsStore.Validate(configured);
            Check(roundtrip.Senders.Length == 3 && roundtrip.Rules[1].Domains[0] == "customer.example", "settings round trip preserves configuration");
            Check(roundtrip.Senders[0].Represented == "", "optional represented sender survives serialization");
            configured.Senders[0].Id = "changed-name";
            Throws(() => SettingsStore.Validate(configured), "renamed sender with dangling rule is rejected");
            var temporary = Path.Combine(Path.GetTempPath(), "autofrom-test-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                SettingsStore.Save(temporary, roundtrip);
                Check(Config.Load(temporary).Enabled, "settings first save creates valid file");
                roundtrip.Enabled = false;
                SettingsStore.Save(temporary, roundtrip);
                Check(!Config.Load(temporary).Enabled && Config.Load(temporary + ".bak").Enabled, "atomic replacement retains prior settings backup");
                string before = File.ReadAllText(temporary);
                roundtrip.Senders[0].Id = "invalid-reference";
                Throws(() => SettingsStore.Save(temporary, roundtrip), "invalid settings cannot be persisted");
                Check(File.ReadAllText(temporary) == before, "invalid save leaves existing settings intact");
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); if (File.Exists(temporary + ".bak")) File.Delete(temporary + ".bak"); }
            fake.FailRestore = false;
            Controller.RestoreManaged(fake, state);
            Check(!state.Managed && fake.Current.Represented == "", "policy change restores managed draft before commit");
            fake.FailRestore = true;
            Controller.RestoreManaged(fake, state);
            Check(!state.Managed, "policy change leaves unmanaged draft untouched");
            var xml = new XmlDocument(); xml.LoadXml(RibbonMarkup.For("Microsoft.Outlook.Explorer"));
            Check(xml.OuterXml.Contains("OpenSettings"), "explorer ribbon XML exposes settings callback");
            xml.LoadXml(RibbonMarkup.For("Microsoft.Outlook.Mail.Compose"));
            Check(RibbonMarkup.For("Microsoft.Outlook.Mail.Read") == null, "ribbon limited to inbox and mail compose");
            var ribbonDefault = (System.Runtime.InteropServices.ComDefaultInterfaceAttribute)Attribute.GetCustomAttribute(typeof(Connect), typeof(System.Runtime.InteropServices.ComDefaultInterfaceAttribute));
            Check(ribbonDefault.Value == typeof(IAutoFromRibbon), "callbacks exposed through default COM dispatch interface");
            using (var form = new SettingsForm(Config.Load(args[1]), new[] { "alex@contoso.com", "alex@fabrikam.com" }, null, c => { }))
            {
                Check(form.ReadConfiguration().Rules.Length == 3, "settings form loads and validates existing rules");
                var tab = Descendants(form).OfType<TabControl>().Single();
                var senderGrid = Descendants(tab.TabPages[0]).OfType<DataGridView>().Single();
                var ruleGrid = Descendants(tab.TabPages[1]).OfType<DataGridView>().Single();
                senderGrid.Rows[0].Cells["id"].Value = "renamed";
                Throws(() => form.ReadConfiguration(), "UI rejects dangling sender names");
                senderGrid.Rows[0].Cells["id"].Value = "personal";
                ruleGrid.Rows[0].Cells["priority"].Value = "not-a-number";
                Throws(() => form.ReadConfiguration(), "UI rejects non-numeric priorities");
                ruleGrid.Rows[0].Cells["priority"].Value = "5";
                ruleGrid.Rows[0].Cells["recipients"].Value = " ONE@Example.com; two@example.com ";
                Check(form.ReadConfiguration().Rules[0].Recipients.SequenceEqual(new[] { "one@example.com", "two@example.com" }), "UI parses recipient lists and normalizes addresses");
            }
            using (var form = new SettingsForm(null, new string[0], "Saved settings are invalid.", c => { }))
                Check(form.ReadConfiguration().Rules.Length == 0, "recovery UI opens with no profile accounts or valid config");
            var aliasConfig = Config.Parse("{\"version\":2,\"enabled\":true,\"senders\":[{\"id\":\"my-alias\",\"account\":\"alex@contoso.com\",\"alias\":\"sales-alias@contoso.com\",\"defaultForAccount\":true}],\"rules\":[]}");
            var aliasTarget = aliasConfig.Senders[0];
            Check(aliasTarget.IsAlias && aliasTarget.FromAddress == "sales-alias@contoso.com", "alias modeled separately from delegated mailbox");
            Check(SettingsStore.Validate(aliasConfig).Senders[0].DefaultForAccount, "alias and default survive persistence");
            Check(RuleEngine.Decide(aliasConfig, "alex@contoso.com", new string[0], true).Target == aliasTarget, "default alias selected on empty compose");
            Check(RuleEngine.Decide(aliasConfig, "alex@contoso.com", new[] { "any@external.example" }, true).Target == aliasTarget, "default alias needs no recipient rule");
            Check(RuleEngine.Decide(aliasConfig, "someone@elsewhere.com", new string[0], true).Target == null, "default does not affect other accounts");
            Check(RuleEngine.Decide(aliasConfig, "alex@contoso.com", new string[0], false).Error != null, "unresolved recipients still block with default");
            var ambiguousDefault = SettingsStore.Validate(aliasConfig);
            ambiguousDefault.Senders = new[] { ambiguousDefault.Senders[0], new Sender { Id = "other", Account = "alex@contoso.com", Alias = "another@contoso.com", Represented = "", DefaultForAccount = true } };
            Throws(() => SettingsStore.Validate(ambiguousDefault), "two defaults for one account rejected");
            var invalidAlias = SettingsStore.Validate(aliasConfig); invalidAlias.Senders[0].Represented = "sales@contoso.com";
            Throws(() => SettingsStore.Validate(invalidAlias), "alias and delegate cannot be mixed");
            invalidAlias.Senders[0].Represented = ""; invalidAlias.Senders[0].Alias = "alex@contoso.com";
            Throws(() => SettingsStore.Validate(invalidAlias), "primary address is not an alias route");
            var aliasSession = new MockSession();
            aliasSession.Accounts.Items[0].Aliases = new[] {"SMTP:alex@contoso.com", "smtp:sales-alias@contoso.com", "x500:ignored"};
            var aliasMail = new MockMail { SendUsingAccount = aliasSession.Accounts.Items[0] };
            var aliasPort = new OutlookPort(aliasMail, aliasSession);
            var originalAliasSnapshot = aliasPort.Capture();
            aliasPort.Set(aliasTarget);
            Check(aliasMail.SentOnBehalfOfName == aliasTarget.Alias && aliasMail.SaveCalls == 1, "alias selection writes literal address and saves draft");
            Check((string)aliasMail.PropertyAccessor.GetProperty(AliasIdentity.SmtpAddress) == aliasTarget.Alias, "SMTP alias is not replaced by primary address");
            Check(aliasSession.ResolveCalls == 0, "alias is never resolved through CreateRecipient");
            aliasPort.Set(aliasTarget);
            Check(aliasMail.SaveCalls == 1, "unchanged alias identity avoids repeated draft saves");
            byte[] entryId = (byte[])aliasMail.PropertyAccessor.GetProperty(AliasIdentity.EntryId);
            Check(entryId.Take(24).SequenceEqual(new byte[] {0,0,0,0,0x81,0x2b,0x1f,0xa4,0xbe,0xa3,0x10,0x19,0x9d,0x6e,0,0xdd,1,0x0f,0x54,2,0,0,1,0x90}), "one-off header matches documented byte order and no-lookup flags");
            Check(System.Text.Encoding.Unicode.GetString(entryId, 24, entryId.Length - 24) == aliasTarget.Alias + "\0SMTP\0" + aliasTarget.Alias + "\0", "one-off entry contains literal UTF16 SMTP alias");
            var aliasSnapshot = aliasPort.Capture();
            aliasPort.Set(new Sender { Account = "alex@contoso.com", Represented = "" });
            Check(aliasMail.SentOnBehalfOfName == "" && aliasMail.PropertyAccessor.Values.Count == 0, "switching to primary clears all alias identity fields");
            aliasPort.Restore(aliasSnapshot);
            Check(AliasIdentity.Matches(aliasPort.Capture().RepresentingProperties, aliasSnapshot.RepresentingProperties), "restoring previously selected alias preserves exact identity");
            aliasPort.Restore(originalAliasSnapshot);
            Check(aliasMail.PropertyAccessor.Values.Count == 0, "original absence of MAPI fields restored");
            aliasSession.Accounts.Items[0].Aliases = new[] {"smtp:someone-else@contoso.com"};
            Throws(() => aliasPort.Set(aliasTarget), "unowned alias rejected before writing draft");
            Check(aliasMail.SentOnBehalfOfName == "", "unowned alias leaves original sender intact");
            aliasSession.Accounts.Items[0].Aliases = new[] {"smtp:alex@contoso.com", "smtp:sales-alias@contoso.com"};
            aliasMail.PropertyAccessor.FailNextSet = true;
            Throws(() => aliasPort.Set(aliasTarget), "MAPI property write failure cancels alias selection");
            Check(aliasMail.SentOnBehalfOfName == "" && aliasMail.PropertyAccessor.Values.Count == 0, "partial alias write rolls back all fields");
            aliasMail.OnSave = m => { m.OnSave = null; m.SentOnBehalfOfName = "Alex Example"; };
            aliasPort.Set(aliasTarget);
            Check(aliasMail.SentOnBehalfOfName == "Alex Example" && AliasIdentity.Matches(aliasPort.Capture().RepresentingProperties, AliasIdentity.For(aliasTarget.Alias)), "display name normalization accepted only with exact alias address metadata");
            int savesBefore = aliasMail.SaveCalls;
            aliasPort.Set(aliasTarget);
            Check(aliasMail.SaveCalls == savesBefore, "normalized display name does not cause repeated saves");
            aliasPort.Restore(originalAliasSnapshot);
            aliasMail.OnSave = m => { m.OnSave = null; m.PropertyAccessor.SetProperty(AliasIdentity.AddressType, ""); m.PropertyAccessor.SetProperty(AliasIdentity.EmailAddress, ""); };
            aliasPort.Set(aliasTarget);
            Check(AliasIdentity.MatchesAlias(aliasPort.Capture().RepresentingProperties, AliasIdentity.For(aliasTarget.Alias)), "observed blank routing pair accepted with exact SMTP and binary identity");
            savesBefore = aliasMail.SaveCalls;
            aliasPort.Set(aliasTarget);
            Check(aliasMail.SaveCalls == savesBefore, "normalized routing pair avoids repeated writes on timer and send");
            var normalized = aliasPort.Capture();
            aliasPort.Restore(originalAliasSnapshot);
            aliasPort.Restore(normalized);
            Check(AliasIdentity.Matches(aliasPort.Capture().RepresentingProperties, normalized.RepresentingProperties), "normalized identity restores exactly");
            foreach (string key in new[] { AliasIdentity.SmtpAddress, AliasIdentity.EntryId, AliasIdentity.SearchKey, AliasIdentity.AddressType, AliasIdentity.EmailAddress })
            {
                var bad = new Dictionary<string, object>(normalized.RepresentingProperties);
                bad[key] = key == AliasIdentity.EntryId || key == AliasIdentity.SearchKey ? (object)new byte[] { 1, 2 } : "wrong";
                Check(!AliasIdentity.MatchesAlias(bad, AliasIdentity.For(aliasTarget.Alias)), "normalized alias rejects conflicting property " + key);
                bad.Remove(key);
                Check(!AliasIdentity.MatchesAlias(bad, AliasIdentity.For(aliasTarget.Alias)), "normalized alias rejects missing property " + key);
            }
            aliasPort.Restore(originalAliasSnapshot);
            aliasMail.OnSave = m => { m.OnSave = null; m.PropertyAccessor.SetProperty(AliasIdentity.SmtpAddress, "alex@contoso.com"); };
            string rewriteReport = "";
            try { aliasPort.Set(aliasTarget); } catch (Exception ex) { rewriteReport = ex.Message; }
            Check(rewriteReport.Contains("0x5D02001F") && rewriteReport.Contains("sales-alias@contoso.com") && rewriteReport.Contains("alex@contoso.com"), "SMTP rewrite rejected and pre-rollback values retained in error");
            Check(aliasMail.SentOnBehalfOfName == "" && aliasMail.PropertyAccessor.Values.Count == 0, "rewrite failure restores previous sender");
            var missingIdentity = AliasIdentity.For(aliasTarget.Alias); missingIdentity.Remove(AliasIdentity.EntryId);
            Check(!AliasIdentity.Matches(missingIdentity, AliasIdentity.For(aliasTarget.Alias)) && AliasIdentity.Differences(missingIdentity, AliasIdentity.For(aliasTarget.Alias)).Contains("<missing>"), "missing identity field rejected and diagnosed");
            aliasSession.Accounts.Items[0].AccountType = 2;
            Throws(() => aliasPort.Set(aliasTarget), "aliases rejected for non-Exchange accounts");
            using (var form = new SettingsForm(aliasConfig, new[] {"alex@contoso.com"}, null, c => { }))
                Check(form.ReadConfiguration().Senders[0].IsAlias && form.ReadConfiguration().Senders[0].DefaultForAccount, "UI loads alias and default checkbox without a rule");
            aliasSession.Accounts.Items[0].AccountType = 0;
            aliasSession.Accounts.Items[0].DirectoryPrimary = "new-primary@contoso.com";
            aliasSession.Accounts.Items[0].Aliases = new[] {"SMTP:new-primary@contoso.com", "smtp:alex@contoso.com", "smtp:sales-alias@contoso.com"};
            aliasPort.Set(aliasTarget);
            Check(aliasMail.SentOnBehalfOfName == aliasTarget.Alias, "account address may differ from directory primary when both belong to same mailbox");
            aliasSession.Accounts.Items[0].Aliases = new[] {"SMTP:new-primary@contoso.com", "smtp:sales-alias@contoso.com"};
            Throws(() => aliasPort.Set(aliasTarget), "target alias alone is insufficient when selected account is not in same mailbox proxy list");
            var eventSession = new MockSession();
            eventSession.Accounts.Items[0].Aliases = new[] { "smtp:alex@contoso.com", "smtp:sales-alias@contoso.com" };
            var connect = new Connect();
            var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
            var eventApp = new MockEventApplication { Session = eventSession };
            typeof(Connect).GetField("application", flags).SetValue(connect, eventApp);
            typeof(Connect).GetField("config", flags).SetValue(connect, aliasConfig);
            var eventMail = new MockMail { SendUsingAccount = eventSession.Accounts.Items[0] };
            var inspector = new MockEventInspector { CurrentItem = eventMail };
            using (var eventTimer = new Timer { Interval = 250 })
            {
            typeof(Connect).GetField("timer", flags).SetValue(connect, eventTimer);
            eventApp.Inspectors.Items = new[] { inspector };
            Action openOnly = () => typeof(Connect).GetMethod("OnNewInspector", flags).Invoke(connect, new object[] { inspector });
            Action tick = () => typeof(Connect).GetMethod("Tick", flags).Invoke(connect, new object[] { null, EventArgs.Empty });
            openOnly();
            Check(eventMail.SaveCalls == 0 && eventTimer.Interval == 50, "compose-open schedules quick retry without writing before UI initialization");
            Action fireOpen = () => { openOnly(); tick(); };
            fireOpen();
            Check(eventMail.SentOnBehalfOfName == aliasTarget.Alias && eventMail.SaveCalls == 1, "deferred compose callback applies default on first UI tick");
            fireOpen();
            Check(eventMail.SaveCalls == 1, "repeated compose callback does not resave matching alias");
            eventMail = new MockMail { SendUsingAccount = eventSession.Accounts.Items[0] }; inspector.CurrentItem = eventMail;
            aliasConfig.Enabled = false; fireOpen();
            Check(eventMail.SaveCalls == 0, "compose callback respects disabled routing");
            aliasConfig.Enabled = true;
            typeof(Connect).GetField("settingsOpen", flags).SetValue(connect, true); fireOpen();
            Check(eventMail.SaveCalls == 0, "compose callback pauses during settings edits");
            typeof(Connect).GetField("settingsOpen", flags).SetValue(connect, false);
            eventMail.Sent = true; fireOpen();
            Check(eventMail.SaveCalls == 0, "compose callback ignores sent messages");
            eventMail.Sent = false; eventMail.SendUsingAccount = null; fireOpen();
            Check(eventMail.SaveCalls == 0, "early compose callback does not guess uninitialized account");
            eventMail.SendUsingAccount = eventSession.Accounts.Items[0]; fireOpen();
            Check(eventMail.SaveCalls == 1, "early compose failure can be retried after account initialization");
                        eventApp.Inspectors.Items = new MockEventInspector[0]; tick();
            var retained = (System.Collections.IDictionary)typeof(Connect).GetField("states", flags).GetValue(connect);
            Check(retained.Count == 0, "discarded draft state cleared after empty scan");
            for (int cycle = 0; cycle < 5; cycle++)
            {
                eventMail = new MockMail { SendUsingAccount = eventSession.Accounts.Items[0] };
                inspector.CurrentItem = eventMail;
                eventApp.Inspectors.Items = new[] { inspector };
                fireOpen();
                Check(eventMail.SaveCalls == 1 && eventMail.SentOnBehalfOfName == aliasTarget.Alias, "new draft after discard applies alias cycle " + cycle);
                eventApp.Inspectors.Items = new MockEventInspector[0]; tick();
            }
            }
            Console.WriteLine(count + " tests passed."); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
    static IEnumerable<Control> Descendants(Control control)
    {
        foreach (Control child in control.Controls)
        {
            yield return child;
            foreach (Control nested in Descendants(child)) yield return nested;
        }
    }
    sealed class Fake : IComposePort
    {
        public Snapshot Current = new Snapshot { Account = "alex@contoso.com", Represented = "" };
        public string[] Addresses = new[] { "one@customer.example" };
        public bool Resolved, FailSet, FailRestore;
        public Snapshot Capture() { return new Snapshot { Account = Current.Account, Represented = Current.Represented }; }
        public IList<string> Recipients(bool resolve, out bool complete) { Resolved = resolve; complete = true; return Addresses; }
        public void Set(Sender target) { Current.Account = target.Account; Current.Represented = target.Represented; if (FailSet) throw new Exception("simulated COM failure"); }
        public void Restore(Snapshot original) { if (FailRestore) throw new Exception("simulated restore failure"); Current = new Snapshot { Account = original.Account, Represented = original.Represented }; }
    }
}

// Public types intentionally exercise the same dynamic member access used for Outlook COM objects.
// These mocks verify the adapter contract; they do not simulate Exchange or COM activation.
public sealed class MockAccount
{
    public string SmtpAddress { get; set; }
    public int AccountType { get; set; }
    public string[] Aliases = new string[0];
    public string DirectoryPrimary;
    public MockRecipient CurrentUser
    {
        get
        {
            var entry = new MockEntry { AddressEntryUserType = 0, Address = DirectoryPrimary ?? SmtpAddress };
            entry.PropertyAccessor.SetProperty(AliasIdentity.ProxyAddresses, Aliases);
            return new MockRecipient { AddressEntry = entry };
        }
    }
}
public sealed class MockAccounts
{
    public MockAccount[] Items = { new MockAccount { SmtpAddress = "alex@contoso.com", AccountType = 0 }, new MockAccount { SmtpAddress = "alex@fabrikam.com", AccountType = 2 } };
    public int Count { get { return Items.Length; } }
    public MockAccount Item(int i) { return Items[i - 1]; }
}
public sealed class MockSession
{
    public MockAccounts Accounts = new MockAccounts();
    public string PrimaryAddress = "sales@contoso.com";
    public int ResolveCalls;
    public MockRecipient CreateRecipient(string text) { ResolveCalls++; return new MockRecipient { AddressEntry = new MockEntry { AddressEntryUserType = 0, Address = PrimaryAddress } }; }
}
public sealed class MockMail
{
    public int Class { get { return 43; } }
    public bool Sent, Submitted;
    public MockAccount SendUsingAccount { get; set; }
    public string SentOnBehalfOfName { get; set; }
    public MockRecipients Recipients = new MockRecipients();
    public MockPropertyAccessor PropertyAccessor = new MockPropertyAccessor();
    public int SaveCalls;
    public Action<MockMail> OnSave;
    public void Save() { SaveCalls++; if (OnSave != null) OnSave(this); }
}
public sealed class MockEventApplication {
    public MockSession Session;
    public MockEventInspectors Inspectors = new MockEventInspectors();
    public MockEventExplorers Explorers = new MockEventExplorers();
}
public sealed class MockEventInspectors {
    public MockEventInspector[] Items = new MockEventInspector[0];
    public int Count { get { return Items.Length; } }
    public MockEventInspector Item(int i) { return Items[i-1]; }
}
public sealed class MockEventExplorers { public int Count { get { return 0; } } }
public sealed class MockEventInspector { public object CurrentItem; }
public sealed class MockRecipients
{
    public MockRecipient[] Items = { new MockRecipient { AddressEntry = new MockEntry { AddressEntryUserType = 30, Address = "one@customer.example" } } };
    public bool ResolveCalled;
    public int Count { get { return Items.Length; } }
    public MockRecipient Item(int i) { return Items[i - 1]; }
    public bool ResolveAll() { ResolveCalled = true; return true; }
}
public sealed class MockRecipient
{
    public bool Resolved { get { return true; } }
    public MockEntry AddressEntry;
    public void Resolve() { }
}
public sealed class MockEntry
{
    public MockPropertyAccessor PropertyAccessor = new MockPropertyAccessor();
    public int AddressEntryUserType;
    public string Address;
    public MockExchangeUser GetExchangeUser() { return new MockExchangeUser { PrimarySmtpAddress = Address }; }
}
public sealed class MockExchangeUser { public string PrimarySmtpAddress; }
public sealed class MockPropertyAccessor
{
    public Dictionary<string, object> Values = new Dictionary<string, object>();
    public bool FailNextSet;
    public object GetProperty(string name)
    {
        object value;
        if (!Values.TryGetValue(name, out value)) throw new System.Runtime.InteropServices.COMException("Property not found", unchecked((int)0x8004010F));
        return value;
    }
    public void SetProperty(string name, object value)
    {
        if (FailNextSet) { FailNextSet = false; throw new System.Runtime.InteropServices.COMException("Property write blocked"); }
        Values[name] = value is byte[] ? ((byte[])value).Clone() : value;
    }
    public void DeleteProperty(string name)
    {
        if (!Values.Remove(name)) throw new System.Runtime.InteropServices.COMException("Property not found", unchecked((int)0x8004010F));
    }
}
