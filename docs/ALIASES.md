# Alias support: version 1.2 pilot

An alias is an extra SMTP address on the **same mailbox**. It does not need a separate mailbox's Send As or Send on Behalf permission. The earlier delegated-mailbox configuration was the wrong route for this scenario.

## Set up an always-used alias

1. Close Outlook, extract the new build, and rerun `install.ps1` (Outlook bitness is now detected automatically). Reinstallation is required because the assembly version changed. Existing settings are retained; open drafts should be reviewed or discarded before upgrading.
2. Open Outlook > AutoFrom > Settings. The title should show **1.2.6 (alias pilot)**.
3. In Senders, use any label you recognize. Under **Owning Outlook account**, select the account that owns the alias.
4. Enter your desired address in **Alias address**. Clear **Shared / delegated mailbox** on this row.
5. Tick **Use as default**. Remove obsolete rules you added only to establish a default. No rule is needed for this mode.
6. Save settings. Start a new email. Its From identity should change on compose-window creation or a subsequent 250 ms fallback tick, even before you add recipients, provided Outlook has assigned an initial account and permits the property writes.
7. Send a test yourself to an external mailbox you control and inspect the received message's full **From** header. Also test internal recipients, replies, inline replies, forwards and reopened drafts. A correct-looking compose field is not sufficient proof of successful delivery from the alias.

The default applies to drafts whose original Outlook account is the owning account. It does not redirect drafts started from other accounts. Explicit recipient rules can override the default; conflicting recipients still block sending. Only one default per account is allowed.

If you already entered the alias in Shared / delegated mailbox, move it to Alias address. The add-in deliberately does not convert existing delegated entries automatically. A label containing an email address does not itself select that address.

`rules.alias.example.json` is a generic developer example. Normal users can configure everything in the ribbon. Saving upgrades configuration to version 2; older add-in builds cannot read version 2. The `.bak` file retains the prior configuration on replacement, but later saves overwrite that backup.

## Exchange Online prerequisites

The alias must actually be assigned to the selected Exchange Online mailbox. Exchange Online's `SendFromAliasEnabled` setting controls whether aliases are preserved instead of rewritten to the primary address. This is an organization-wide administrator setting, not something the add-in changes. [Microsoft cmdlet documentation](https://learn.microsoft.com/powershell/module/exchange/set-organizationconfig?view=exchange-ps#-sendfromaliasenabled).

An Exchange administrator can inspect it in an existing Exchange Online PowerShell session:

```powershell
Get-OrganizationConfig | Select-Object SendFromAliasEnabled
```

If needed, the administrator can enable it after considering the tenant-wide effect:

```powershell
Set-OrganizationConfig -SendFromAliasEnabled $true
```

Microsoft's Exchange team documents Windows Outlook alias support and a fix for online-mode/missing-OAB cases in Version 2403, Build 17425.20236. The same article lists desktop shared-mailbox alias sending as unsupported. This pilot therefore targets aliases belonging to the selected user's Exchange Online account, not aliases of delegated/shared mailboxes, POP/IMAP accounts, or on-premises Exchange. [Exchange team alias-support article](https://techcommunity.microsoft.com/t5/exchange-team-blog/sending-from-email-aliases-public-preview/ba-p/3070501).

## Implementation and evidence boundary

This is an implementation using documented Outlook/MAPI primitives. Microsoft does **not** document a dedicated Outlook Object Model `SelectAlias` method or certify this exact sequence as an alias-selection SDK. The sequence is an engineering implementation that needs live client/tenant testing; the automated suite does not prove alias delivery or immediate ribbon From-menu refresh.

Ownership is checked against the selected Account's CurrentUser address entry and `PidTagAddressBookProxyAddresses` (0x800F101F). Both the configured account address and requested alias must appear in that same proxy list. The configured account address need not equal the directory primary address. An alias resolving to the same primary mailbox is not rejected; it is verified as an alternate SMTP address on that account. Missing or unavailable proxy metadata blocks alias selection rather than guessing ownership. [Proxy address specification](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxoabk/fd633c11-37bc-4f3c-8165-2ed16686ad3b).

The alias route sets `SendUsingAccount` and the literal represented name. Through `PropertyAccessor`, it sets the represented SMTP address, email address, address type, entry ID, and search key consistently. It uses a Unicode SMTP one-off entry with no directory lookup to avoid normalizing the alias into a primary directory address. It does not change the authenticated sender (`PR_SENDER_*`), transport headers, tenant permissions, or recipients. [PropertyAccessor writes](https://learn.microsoft.com/en-us/office/vba/api/outlook.propertyaccessor.setproperty), [represented SMTP address](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagsentrepresentingsmtpaddress-canonical-property), [one-off entry format](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcdata/b32d23af-85f6-4e92-8387-53a1950ae7ba).

Sender changes explicitly save the current draft, as required for persisting PropertyAccessor changes. This can create a saved blank draft when a default is applied. The add-in never submits a message itself. It compares the exact account and represented identity after saving and again through the ordinary send-time evaluation. A matching unchanged identity avoids repeated saves. Alias write failures roll back all captured identity properties, including deleting properties that were originally absent; rollback failure is reported and sending remains blocked.

Local verification catches observable client-side rewrites and blocked writes. It cannot prove the tenant setting, enforce the server's final behavior after submission, or prevent another add-in from modifying a message afterward. Always verify the delivered From header before relying on this pilot. No messages were sent and no tenant settings were changed during development.

## If From does not change

After saving settings and closing the settings window, open a new email, wait two seconds, and click AutoFrom > Check draft in that email window. Copy the report using Ctrl+A, Ctrl+C. This read-only report shows whether routing matched, whether the timer evaluated this draft, any last automatic-selection error, and the current account/represented sender. It does not resolve recipients, change the draft, or send mail. Review the report before sharing because it contains account addresses. A retained alias property is not proof the visible From field or delivered message uses the alias.

Version 1.2.3 permits Outlook to normalize the represented display name. Version 1.2.4 additionally accepts the observed Save behavior where both address-type and email-address fields become empty strings, provided the owning account, represented SMTP address, full one-off entry ID, and search key still match exactly. Missing fields, only one blank field, and conflicting values remain rejected. This is a compatibility decision based on a live diagnostic report, not a Microsoft guarantee of alias delivery. A display name is not itself an SMTP identity. See [Microsoft's represented-name definition](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/34c5170c-5c81-4a85-8cc8-31ae98d7f58c). Rejected writes now report expected and read-back values before rollback; the current sender shown later in the report is the restored identity. This distinction avoids misreading rollback as proof that Outlook removed every field.

Version 1.2.5 uses [NewInspector](https://learn.microsoft.com/en-us/office/vba/api/outlook.inspectors.newinspector), which Microsoft documents as occurring before the window appears. The handler applies selection immediately if the draft is ready. Inline replies and drafts without an initialized account use the 250 ms fallback scan. This interval is not a guaranteed latency: Outlook owns the UI thread, initialization, and saving. There is no guarantee that the primary address never flashes. Check draft reports whether the compose-open event hookup succeeded.

Version 1.2.6 defers the initial sender write to a 50 ms Windows Forms timer tick after NewInspector, then returns to 250 ms scanning. It avoids the pre-display write introduced in 1.2.5, which can leave visible From stale while the underlying alias is correct. This addresses the suspected timing cause; live UI refresh after repeated discard/reopen cycles still needs verification. Timer intervals are not guaranteed wall-clock latency.
