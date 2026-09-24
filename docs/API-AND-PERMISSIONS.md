# API findings and mailbox permissions

Checked against Microsoft documentation on September 23, 2026.

## Supported approaches

| Client / technology | Can change native compose From? | Project support |
| --- | --- | --- |
| Classic Outlook on Windows, COM/Outlook Object Model | Yes: choose a configured sending account and set a represented sender | Implemented; live pilot validation required |
| Classic Outlook, Office.js web add-in | No documented From setter | Not implemented |
| New Outlook on Windows | Office.js lacks the setter; COM/VSTO is unsupported | Not supported |
| Outlook on the web | Office.js lacks the setter | Not supported |
| Outlook on Mac/mobile | No From setter through the documented Office.js interface; this Windows COM DLL cannot load | Not supported |

[`Office.From`](https://learn.microsoft.com/en-us/javascript/api/outlook/office.from?view=outlook-js-preview) exposes only `getAsync` in compose mode, introduced in Mailbox 1.7 with ReadItem minimum permission. Even the documented preview interface has no `setAsync`. Requesting ReadWriteMailbox or using a shared mailbox does not add a From-writing method. This applies to an Office.js add-in even when classic Outlook is its host.

Compose/recipient/from/send events can activate an add-in, but do not introduce a setter. Relevant events include `OnNewMessageCompose` (1.10), `OnMessageRecipientsChanged` (1.11), `OnMessageSend` (1.12), and `OnMessageFromChanged` (1.13). The last observes a sender change; it cannot perform one. Client/build/channel support is event-specific. [Microsoft event-based activation documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation).

Microsoft explicitly excludes COM and VSTO from new Outlook while retaining them in classic Outlook. This is why an Office.js manifest or a new Outlook “sideload” cannot install this solution. [Develop add-ins for new Outlook](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/one-outlook).

For classic Outlook, [`MailItem.SendUsingAccount`](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.sendusingaccount) is a writable Account reference used for submission. [`MailItem.SentOnBehalfOfName`](https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.sentonbehalfofname) is the writable represented-sender name property. This project supplies the resolved primary SMTP address, using Outlook's ordinary delegated-send path. It does not forge transport headers or change the read-only actual sender property. Exchange decides the final Send As versus on-behalf identity.

[`Application.ItemSend`](https://learn.microsoft.com/en-us/office/vba/api/outlook.application.itemsend) permits cancellation before submission, including when Outlook's Send method initiates sending. The final check uses this event. [`Explorer.ActiveInlineResponse`](https://learn.microsoft.com/en-us/office/vba/api/outlook.explorer.activeinlineresponse) exposes inline reply items for property editing; AutoFrom does not invoke its unsupported Send method.

## Shared mailboxes and delegated permissions

Full Access permits mailbox access, **not** sending. Send As makes mail appear from the target mailbox; Send on Behalf identifies the delegate acting for it. These are Exchange permissions, independent of any add-in permission. When both sending rights exist, Send As takes precedence. The property name `SentOnBehalfOfName` therefore is not an instruction to override Exchange permission semantics. Configure and inspect received test messages to establish the actual result. [Manage permissions in Exchange Online](https://learn.microsoft.com/en-us/exchange/recipients-in-exchange-online/manage-permissions-for-recipients).

An Exchange administrator can use Mailbox delegation in the Exchange admin center. The following are reference commands for a connected Exchange Online PowerShell administrator; the project does **not** execute them. Substitute real addresses and grant only the needed permission:

```powershell
# Choose this for Send As:
Add-RecipientPermission -Identity sales@contoso.com -Trustee alex@contoso.com -AccessRights SendAs

# OR use this for Send on Behalf (preserves existing delegates):
Set-Mailbox -Identity sales@contoso.com -GrantSendOnBehalfTo @{Add='alex@contoso.com'}

# Only if the user also needs to open/manage the shared mailbox:
Add-MailboxPermission -Identity sales@contoso.com -User alex@contoso.com -AccessRights FullAccess -InheritanceType All
```

Permissions may take time to propagate. A mailbox hidden from address lists can prevent desktop delegated sending; the target must be resolvable and visible. [Microsoft 365 mailbox permission guidance](https://learn.microsoft.com/en-us/microsoft-365/admin/add-users/give-mailbox-permissions-to-another-user?view=o365-worldwide).

For comparison, Office.js shared-folder support starts at Mailbox 1.8, and shared-mailbox support at 1.13. Supporting those contexts requires the relevant shared-folder manifest capability (`SupportsSharedFolders` in the add-in-only manifest). Compose `getSharedPropertiesAsync` also has access-context restrictions; opening another mailbox in its own web window or promoting it to a full account in new Outlook changes those conditions. These capabilities expose shared contexts and do not permit setting From. [Shared folder and shared mailbox add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/delegate-access).

An automatically attached shared folder is not necessarily a sending Account in classic Outlook. The implemented delegated route explicitly selects the user's configured Exchange transport account and the shared mailbox as represented sender. The delegated route restricts targets to resolvable Exchange mailbox user entries. Distribution groups are outside its scope. Version 1.2 adds a separate alias route for proxy addresses on the selected account; it does not reuse the delegated-address resolver. See [alias implementation and prerequisites](ALIASES.md) for the MAPI approach and its live-validation limits.

## Why not Microsoft Graph or a browser extension?

Graph can send from another mailbox. Delegated token use generally requires `Mail.Send.Shared` plus the relevant Exchange sending permission; sending through the other user's mailbox endpoint also needs Full Access in the documented delegated scenario. This is a separate send workflow, not an API for manipulating the native Outlook compose From selector. [Microsoft Graph sending as another user](https://learn.microsoft.com/en-us/graph/outlook-send-mail-from-other-user).

Replacing Outlook's Send with Graph would require a separately designed composer or interception flow, authentication, and testing for attachments, inline images, signatures, encryption/labels, threading, saved drafts, Sent Items, and duplicate prevention. Patching a server-side draft while Outlook is editing it is not a documented synchronization contract for this UI. No such workaround is used here.

A browser content script could attempt to click the From UI in Outlook on the web. That would depend on private DOM behavior, localization, and timing, and would not provide a supported implementation for the new Outlook desktop client. This project uses the supported classic Outlook Object Model instead.

## Local add-in permissions and operational limits

There is no Entra registration, Graph consent, web-add-in permission scope, or server service. A native in-process COM add-in has broad user-level access; this implementation narrows what it actually reads and changes, but it is not sandboxed. It uses only the Application instance supplied by Outlook, following Microsoft's [Object Model security guidance](https://learn.microsoft.com/en-us/office/client-developer/outlook/selecting-an-api-or-technology-for-developing-solutions-for-outlook).

Local setter verification is not a permission check or delivery guarantee. Offline use, GAL availability, permission propagation, other add-ins, profile configuration, client disablement, and server policies affect the result. Sent Items location remains controlled by Outlook/Exchange configuration. No arbitrary SMTP spoofing or alias-sending guarantee is provided. If new Outlook/web is mandatory, the exact native-compose requirement currently has no supported implementation in the documented add-in API; a separately scoped custom composer would be necessary.
