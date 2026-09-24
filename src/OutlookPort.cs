using System;
using System.Collections.Generic;

namespace AutoFrom
{
    public sealed class OutlookPort : IComposePort
    {
        readonly dynamic mail, session;
        public OutlookPort(object item, object outlookSession) { mail = item; session = outlookSession; }
        public Snapshot Capture()
        {
            dynamic account = mail.SendUsingAccount;
            // Do not guess a default account; this is also needed to undo automatic changes safely.
            if (account == null) throw new InvalidOperationException("Choose an initial From account in Outlook before using AutoFrom.");
            return new Snapshot { AccountObject = account, Account = Config.Address((string)account.SmtpAddress),
                Represented = (string)mail.SentOnBehalfOfName ?? "",
                RepresentingProperties = AliasIdentity.Capture((object)mail.PropertyAccessor) };
        }
        public IList<string> Recipients(bool resolve, out bool complete)
        {
            dynamic recipients = mail.Recipients;
            complete = true;
            if (resolve && !(bool)recipients.ResolveAll()) complete = false;
            var addresses = new List<string>();
            for (int i = 1; i <= (int)recipients.Count; i++)
            {
                dynamic r = recipients.Item(i);
                if (!(bool)r.Resolved) { complete = false; continue; }
                try { addresses.Add(Smtp(r.AddressEntry, false)); }
                catch { complete = false; }
            }
            return addresses;
        }
        static string Smtp(dynamic entry, bool requireExchange)
        {
            int kind = (int)entry.AddressEntryUserType;
            // Exchange users, remote Exchange users, and individual SMTP entries only.
            // Reject Exchange distribution lists and local contact groups: never guess membership.
            if (kind == 0 || kind == 5)
            {
                dynamic user = entry.GetExchangeUser();
                if (user == null) throw new InvalidOperationException("Cannot resolve Exchange user.");
                return Config.Address((string)user.PrimarySmtpAddress);
            }
            if (!requireExchange && kind == 30) return Config.Address((string)entry.Address);
            throw new InvalidOperationException("Only individual SMTP recipients and Exchange mailbox senders are supported.");
        }
        string ResolveRepresented(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            dynamic r = session.CreateRecipient(value);
            r.Resolve();
            if (!(bool)r.Resolved) throw new InvalidOperationException("Delegated sender could not be resolved in the address book.");
            return Smtp(r.AddressEntry, true);
        }
        object FindAccount(string address)
        {
            object found = null;
            dynamic accounts = session.Accounts;
            for (int i = 1; i <= (int)accounts.Count; i++)
            {
                dynamic a = accounts.Item(i);
                if (string.Equals((string)a.SmtpAddress, address, StringComparison.OrdinalIgnoreCase))
                {
                    if (found != null) throw new InvalidOperationException("Multiple profile accounts use the configured SMTP address.");
                    found = a;
                }
            }
            if (found == null) throw new InvalidOperationException("The configured sending account is not in this Outlook profile.");
            return found;
        }
        public void Set(Sender target)
        {
            dynamic account = FindAccount(target.Account);
            if (target.IsAlias) { SetAlias(target, (object)account); return; }
            if (target.Represented.Length > 0)
            {
                if ((int)account.AccountType != 0) throw new InvalidOperationException("Delegated sending requires an Exchange account.");
                if (ResolveRepresented(target.Represented) != target.Represented)
                    throw new InvalidOperationException("This is an alias. Move it to the Alias column and clear Shared / delegated mailbox.");
            }
            var before = Capture();
            // A matching display name alone is insufficient after an alias selection.
            object previousSmtp;
            bool literalAlias = before.RepresentingProperties.TryGetValue(AliasIdentity.SmtpAddress, out previousSmtp) &&
                !string.Equals(Convert.ToString(previousSmtp), target.FromAddress, StringComparison.OrdinalIgnoreCase);
            if (!literalAlias && before.Account == target.Account && ResolveRepresented(before.Represented) == target.Represented) return;
            try
            {
                mail.SendUsingAccount = account;
                AliasIdentity.Restore((object)mail.PropertyAccessor, new Dictionary<string, object>());
                mail.SentOnBehalfOfName = target.Represented;
                mail.Save();
                var after = Capture();
                if (after.Account != target.Account || ResolveRepresented(after.Represented) != target.Represented)
                    throw new InvalidOperationException("Outlook did not retain the selected sender.");
            }
            catch
            {
                try { Restore(before); } catch { /* Caller cancels send even when rollback fails. */ }
                throw;
            }
        }
        void SetAlias(Sender target, object account)
        {
            AliasIdentity.ValidateOwnership(account, target.Account, target.Alias);
            var expected = AliasIdentity.For(target.Alias);
            var before = Capture();
            // SentOnBehalfOfName is a display name, not an SMTP identity assertion.
            // Require the owning account and the verified alias identity instead.
            if (before.Account == target.Account &&
                AliasIdentity.MatchesAlias(before.RepresentingProperties, expected)) return;
            try
            {
                mail.SendUsingAccount = (dynamic)account;
                mail.SentOnBehalfOfName = target.Alias;
                AliasIdentity.Restore((object)mail.PropertyAccessor, expected);
                mail.Save();
                var after = Capture();
                if (after.Account != target.Account ||
                    !AliasIdentity.MatchesAlias(after.RepresentingProperties, expected))
                    throw new InvalidOperationException("Outlook changed the alias identity after saving. Values captured BEFORE restoring the original sender:\n" +
                        "Expected account: " + target.Account + "; read back account: " + after.Account + "\n" +
                        "Read back display name: " + after.Represented + "\n" +
                        AliasIdentity.Differences(after.RepresentingProperties, expected));
            }
            catch (Exception ex)
            {
                try { Restore(before); }
                catch (Exception rollback) { throw new InvalidOperationException("Alias selection and sender restoration both failed. Do not send this draft until its From address is corrected. " + rollback.Message, ex); }
                throw new InvalidOperationException("Could not select the alias automatically. " + ex.Message, ex);
            }
        }
        public void Restore(Snapshot original)
        {
            mail.SendUsingAccount = (dynamic)original.AccountObject;
            mail.SentOnBehalfOfName = original.Represented;
            if (original.RepresentingProperties != null) AliasIdentity.Restore((object)mail.PropertyAccessor, original.RepresentingProperties);
            mail.Save();
            var current = Capture();
            if (current.Account != original.Account || !string.Equals(current.Represented, original.Represented, StringComparison.OrdinalIgnoreCase) ||
                (original.RepresentingProperties != null && !AliasIdentity.Matches(current.RepresentingProperties, original.RepresentingProperties)))
                throw new InvalidOperationException("Could not restore the original sender. Set From manually and reopen the draft.");
        }
    }
}
