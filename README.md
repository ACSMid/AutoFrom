# AutoFrom for classic Outlook on Windows

AutoFrom automatically selects the outgoing sender using a per-account default or recipient/domain/account rules. It supports account and delegated-mailbox routes, with an **alias-aware pilot implementation (current version: 1.2.6)**. You still press Outlook's normal Send button.

**Platform limit:** this project requires classic Outlook for Windows. Microsoft's current Office.js `From` interface has a getter and no setter, and new Outlook does not load COM add-ins. It therefore cannot change the native compose From field in new Outlook or Outlook on the web. See [API findings and sources](docs/API-AND-PERMISSIONS.md).

Choose **AutoFrom > Settings** in Outlook's main or compose ribbon. Version 1.2 adds a separate **Alias address** field and **Use as default** checkbox. You do not need to invent a recipient rule just to use one address. See [alias setup and limitations](docs/ALIASES.md).

The repository includes source, tests, and setup scripts; the installable ZIP includes compiled development binaries. Compilation and 113 automated tests have passed. Alias selection and repeated compose/discard behavior were reported working on one live classic Outlook profile. This is limited user validation, not a compatibility guarantee across clients or tenants. Confirm the delivered From header on your own mailbox. This is a pilot build, not a signed public-release installer. See [GitHub release preparation](docs/RELEASING.md).

## Quick start / local sideload

1. Use Windows with **classic Outlook**, a configured mail profile, and .NET Framework 4.8. In Outlook, File > Office Account > About Outlook shows 32-bit or 64-bit. This refers to **Outlook's** bitness, not Windows's.
2. Extract the entire project to a local folder. If using a source-only download, open Windows PowerShell in that folder and build first. The installable ZIP already includes compiled binaries:

   ```powershell
   powershell.exe -NoProfile -File .\build.ps1
   ```

3. Close classic Outlook completely. Double-click **Install AutoFrom.cmd**, or run:

   ```powershell
   powershell.exe -NoProfile -File .\install.ps1
   ```

4. Start classic Outlook. Check File > Options > Add-ins > Manage COM Add-ins > Go for **AutoFrom - automatic sender rules**. Open the **AutoFrom** ribbon tab and click **Settings**.
5. On **Senders**, click **Add sender**, give it a unique label, and select its **Owning Outlook account**. For an alias on that account, enter **Alias address** and leave Shared / delegated mailbox blank. For a different delegated mailbox, use its primary address in **Shared / delegated mailbox** and leave Alias blank. Leave both blank to use the account itself.
6. For the same sender on all drafts from this account, tick **Use as default**. No rule is required. For optional recipient-specific overrides, use **Rules > Add rule**, select a sender, and enter recipients or domains. Separate entries with semicolons. Original accounts optionally restrict when a rule applies. Lower priority numbers win per recipient.
7. On **Test rules**, enter an original account and all recipients, then click **Test selection**. This previews unsaved rules locally and does not check Exchange permissions or send mail.
8. Click **Save settings**. Changes apply immediately, without restarting Outlook. The shipped configuration starts empty; add your own senders and rules before expecting automatic changes. Cancel discards unsaved edits.
9. Compose a new draft and show its From field via Options > From. A default can apply before recipients are added; recipient rules require resolved recipients. Selection is scheduled 50 ms after a compose window opens, with a 250 ms fallback retry when Outlook is not ready; a brief primary-address flash can still occur. For an alias, verify a delivered message's From header, not just the compose label. Follow [the live acceptance checklist](docs/TESTING.md) before normal use.

No web manifest, web server, Entra application, API key, OAuth grant, or Microsoft 365 web-add-in sideload is involved. The installer registers a COM DLL under HKCU and copies it to `%LOCALAPPDATA%\AutoFrom`; it does not change Exchange permissions or require elevation by design. Enterprise policy may prohibit user-installed or unsigned COM add-ins. Have IT package/sign the reviewed build when required; do not disable organizational security controls.

## Rule behavior

Each allowed sender has an `id` and a profile `account` SMTP address. An optional `represented` primary SMTP address selects a delegated/shared mailbox through that Exchange account. An auto-mapped shared mailbox is not necessarily an entry in Outlook's Accounts collection; normally use your personal Exchange account as `account` and the shared mailbox as `represented`. If the shared mailbox is a genuine configured sending account, it can instead be an account-only route after live verification.

For an alias, use `alias` instead of `represented`. It must be a proxy address belonging to the selected account. `defaultForAccount: true` uses that sender for its owning account's drafts when no explicit recipient rule matches, including an empty draft. Defaults are per original account, not a global override of unrelated accounts. Multiple defaults for one account are rejected.

Each rule has `id`, integer `priority`, and `sender`. Optional conditions:

| Field | Meaning |
| --- | --- |
| `recipients` | Exact SMTP address matches |
| `domains` | Exact recipient domains |
| `includeSubdomains` | Opt-in to matching child domains; default false |
| `accounts` | Limit the rule to the draft's original sending account |

For a rule with recipient and domain lists, either list may match. An account condition, if present, must also match. A rule containing only `accounts` is an account fallback. At least one condition is required. Lower `priority` wins **per recipient**. Exact addresses are not inherently higher priority than domains: assign the priorities explicitly. Rule order does not matter. Address comparison is case-insensitive and uses resolved SMTP addresses, which can differ from typed Exchange aliases.

Every To, Cc, and Bcc recipient must select the same sender. Equal-priority conflicting rules, different senders for different recipients, or a mixture of matching and unmatched recipients blocks sending. This means a sales-domain rule plus a personal fallback will block a message containing both sales customers and unrelated recipients. Add explicit same-sender rules for legitimate combinations or split the message; a more specific rule never silently controls unrelated recipients.

No matching rule or default leaves the draft unchanged. If AutoFrom already changed that open draft and all applicable selections disappear, it restores the original sender. An empty draft can use its account default. Unresolved recipients and unsupported address entries are deferred while typing and block Send when routing is active. Exchange distribution lists and Outlook contact groups must be expanded to individual recipients; membership is never guessed. This version accepts Exchange user entries and plain SMTP entries, not local Outlook Contact address-entry objects, fax, or arbitrary directory types. Re-enter unsupported contacts as plain SMTP addresses.

Rules are authoritative while active: manually changing From on a matched draft is overwritten at the next successful evaluation, including at Send. There is no one-message bypass. To disable globally, uncheck **Automatically select the sender** in Settings and save. Saving first restores original senders on currently open drafts managed by this add-in, then applies the new settings. A failed restoration or file write keeps the previous policy active and shows an error. If you instead disable/uninstall the COM add-in directly, it cannot restore drafts; review their From fields yourself.

The settings editor stores configuration at `%LOCALAPPDATA%\AutoFrom\rules.json` using an atomic file replacement and retains the previous file as `rules.json.bak`. Editing this file is optional for developers. External file edits require an Outlook restart; edits made through Settings do not. Sender names are rule references: after renaming a sender, update any rules that use it. Validation rejects missing or duplicate names, invalid addresses, and invalid priorities before saving. The Test rules page can expose conflicts for a specific recipient combination; it does not prove that all possible combinations are conflict-free.

The original account is captured when a draft is first evaluated in its open session and retained while that draft is open. Closing/reopening a saved draft, popping an inline response into another window if Outlook creates a different item wrapper, or restarting Outlook establishes a new baseline from the draft's then-current account. If you need stable account scoping across those transitions, prefer recipient/domain rules without `accounts` and verify saved drafts. There are no hidden persisted message properties.

## Runtime and safety boundaries

- The compose-open event schedules sender selection after a 50 ms UI-loop delay; a 250 ms Outlook UI-thread fallback timer examines open, unsent compose windows and inline responses. It does not enumerate mailbox folders or read message bodies/attachments. Account changes may not refresh signatures; inspect the resulting signature yourself.
- Version 1.2 saves changed sender properties to the draft and verifies them afterward. A default can therefore create a saved blank draft. Unchanged alias identity does not trigger repeated saves. Alias writes and restoration include represented-sender MAPI metadata; authenticated-sender properties are not changed. See ALIASES.md for the implementation's evidence boundary.
- The normal Outlook `ItemSend` event resolves recipients, evaluates again, applies the chosen sender, and verifies the local properties. A failure cancels that send attempt. The code never calls `Send`, performs a Graph send, or creates a duplicate message.
- An invalid/missing configuration blocks mail sending while the loaded add-in is active and shows an explanation. Open AutoFrom > Settings to enter a valid configuration and save; the previous file is backed up. Settings editing pauses automatic changes, and any send attempted while that window is open is canceled. Configuration is loaded at startup and when saved through the editor.
- Outlook/Exchange remains responsible for authentication, delegated permissions, delivery, and Sent Items. Successful property assignment does not prove send rights; Exchange may reject the message later with an NDR.
- Outlook can disable an add-in or another add-in can change a sender after this add-in's callback. Checks only apply while AutoFrom is loaded and executing; this is not a server-enforced compliance control. Other clients, scheduled sends, background systems, and protected/custom forms require separate validation.
- No telemetry, service calls, credentials, bodies, or recipient logs are produced. Outlook's own address resolution may access your directory/server. Like other in-process COM add-ins, this DLL is not sandboxed and runs with the user's privileges.

## Update, remove, and troubleshoot

Close Outlook before changing binaries. Rebuild and rerun the installer (bitness is detected automatically); existing rules stay in place. When upgrading from version 1.0 or 1.1, rerun the installer rather than only copying the DLL, because its assembly version changed. Move any mistakenly configured aliases from Shared / delegated mailbox to Alias address in the editor. Settings changes through the ribbon do not require a restart. If your Office bitness changes, run `uninstall.ps1` before reinstalling for the new bitness.

To uninstall, close Outlook and run:

```powershell
powershell.exe -NoProfile -File .\uninstall.ps1
```

You can also double-click **Uninstall AutoFrom.cmd**, or use Windows Installed Apps after installing this version. The launcher keeps errors visible. The uninstaller checks both registry views, removes and verifies AutoFrom current-user registrations and its installed DLL, and preserves settings by default. Use `-RemoveSettings` to remove settings and their backup too. Use `-CheckOnly` for a read-only diagnosis or `-WhatIf` for a preview. Outlook must be fully closed. A separate machine-wide registration is reported and must be removed by its administrator/installer. Restart Outlook to clear the cached ribbon. Already changed drafts retain their sender until edited.

If the add-in is missing, check Outlook bitness, COM Add-ins and Disabled Items, the installed DLL path, .NET Framework, and your organization's add-in policies. If a script is blocked, ask IT for its approved signing/deployment process. If a target account is missing, add it through Outlook or choose the delegated route. If the sender cannot resolve, use a visible Exchange mailbox's primary SMTP address. If a From account is initially null, select a real account in Outlook so AutoFrom can capture and restore it safely.

## Project files

`src/Rules.cs` validates configuration and makes deterministic routing decisions. `src/Controller.cs` retains/restores original senders. `src/OutlookPort.cs` implements Outlook address resolution and sender writes. `src/Connect.cs` supplies COM lifecycle, timer, send interception, and live policy updates. `src/Ribbon.cs` exposes the Outlook ribbon and COM callbacks. `src/SettingsForm.cs` is the Windows Forms editor. `src/SettingsStore.cs` validates, serializes, and atomically saves settings. `AutoFrom.csproj` is a .NET Framework project; `build.ps1` uses the Framework compiler with no NuGet dependencies. `tests/Tests.cs` contains executable tests. `bin/` contains development binaries, not an installed add-in.

Installer detection reads the actual classic OUTLOOK.EXE PE architecture after locating it through App Paths, Office InstallRoot, and standard Office folders. It does not infer Outlook bitness from Windows. Missing or conflicting installations stop before registration changes. Advanced override: -OutlookBitness x86 (32-bit) or -OutlookBitness x64 (64-bit). New Outlook alone is unsupported.

Version 1.2.6 defers the initial sender write to a 50 ms Windows Forms timer tick after NewInspector, then returns to 250 ms scanning. It avoids the pre-display write introduced in 1.2.5, which can leave visible From stale while the underlying alias is correct. This addresses the suspected timing cause; live UI refresh after repeated discard/reopen cycles still needs verification. Timer intervals are not guaranteed wall-clock latency.

## License

AutoFrom is source available under the [PolyForm Noncommercial License 1.0.0](LICENSE.md). The license permits noncommercial use, modification, and redistribution, with express provisions for the organizations listed in its text. It does not grant commercial-use rights. Commercial use requires separate permission from the rights holder. Because use is restricted, this is not an open-source license. The full license controls.

## Privacy

The shipped configuration is empty. Examples and tests use fictional addresses. Installed account mappings stay in the current Windows user's local settings file and are not included in this repository. Do not post real settings, diagnostic reports, screenshots, or recipient lists in public issues without redacting them.
