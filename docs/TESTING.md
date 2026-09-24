# Validation and acceptance tests

## Automated checks performed

The project was compiled on Windows with the installed .NET Framework C# compiler, AnyCPU and warnings as errors. **113 automated tests passed.** `build.ps1` compiles both the DLL and the executable test suite, then runs the tests. The rule/controller tests cover exact address selection, domain boundaries, subdomain opt-in, priority, mixed recipients, ambiguous rules, unmatched recipients, original-account conditions, disabling, sender restoration, manual override correction, partial write failures, restoration failures, and configuration errors. Adapter tests use public mock objects to exercise the same dynamic member access for account/delegated writes, snapshot restoration, recipient resolution, missing accounts, non-Exchange delegation, aliases, groups, and null initial accounts.

Version 1.1 also tests configuration serialization, atomic file replacement/backup, validation failure preserving the old file, managed-draft restoration before policy changes, valid ribbon XML, default COM dispatch metadata, editor load/readback, invalid priorities, dangling sender names, address-list normalization, and opening the recovery editor without valid settings. These tests construct Windows Forms controls without showing an Outlook window. They do not verify ribbon callback invocation inside Outlook or replace live UI/DPI testing.

The ApplicationEvents_11 interface GUID and ItemSend DISPID were independently checked against the installed Microsoft Outlook type library without starting Outlook: GUID `0006302c-0000-0000-c000-000000000046`, DISPID `61442`, two event parameters.

Version 1.2 adds alias/default serialization, selection without recipients or rules, per-account isolation, multiple-default rejection, alias/delegate separation, ownership checks, literal SMTP property writes, one-off byte layout, exact identity restoration, partial-write rollback, no redundant saves, and rejection of an observable rewrite after Save. MAPI calls and Exchange identities in these tests are mocks. They do not prove that the live host accepts the sequence or that Exchange preserves the alias after submission.

Compilation and unit tests do not prove COM activation, UI updates, Exchange permissions, or delivery. The user has reported opening an earlier build's ribbon and settings. The agent has not run the installer against the user's registry, opened an Outlook profile, or sent messages. Version 1.2 still requires live alias validation.

## Required live pilot tests

Use a controlled test profile and test recipients. For any test involving actual delivery, the tester presses Send; the test suite never sends mail.

| Test | Expected result |
| --- | --- |
| Load on installed classic Outlook bitness | AutoFrom listed/enabled under COM Add-ins; no startup errors |
| Main-window and compose ribbons | AutoFrom tab appears; Settings opens a single modal editor |
| Sender account list | Existing Outlook account SMTP addresses appear; delegated mailbox remains an optional separate field |
| Add/edit/remove senders and rules | Tables remain usable when switching tabs; invalid fields or deleted sender references prevent Save |
| Test rules page | Unsaved rules can be tested, with no draft creation or mail delivery |
| Cancel or close Settings | Unsaved edits discarded; previous rules resume |
| Save changed rules | Existing managed drafts restored first, settings backed up, new rules apply immediately |
| Disable using Settings | Managed open drafts restored, subsequent automatic selection stops |
| Failed restore or read-only configuration file | Error shown; previous configuration stays active; no misleading successful save |
| High DPI and keyboard navigation | All controls readable and reachable; test 100%, 150%, and 200% scaling |
| Empty shipped config | Normal composition unchanged |
| Alias set as account default, no rules | New draft From changes before recipients are entered; literal alias identity survives a draft save |
| Alias entered under delegated mailbox instead | Actionable message directs user to the Alias column |
| Rule references a deleted/renamed sender label | Error identifies the rule and missing label; removing the unneeded rule permits default-only setup |
| Alias not owned by selected account / proxy metadata unavailable | Settings save or sender selection fails with an explanation; no guessed identity |
| Alias default with other accounts open | Only drafts belonging to the configured owning account are affected |
| Alias changed to another alias / back to primary | Old identity fields are replaced or cleared consistently |
| Alias property blocked/read-only or rewritten during Save | Selection fails; prior identity restored if possible; sending canceled |
| Alias delivery with tenant alias setting enabled | Inspect received From headers at both internal and external test mailboxes; verify alias, reply address, and Sent Items behavior |
| Alias delivery with tenant alias setting disabled | Server rewrite/rejection may occur after local verification; this is not a successful deployment test |
| New draft, exact recipient | Visible From becomes configured profile account within about one second after recipient resolution |
| Domain mapping | From becomes delegated mailbox; near-match domains do not match |
| Reply, reply-all, forward in a window | Same routing behavior, original body/thread preserved |
| Inline reply and then pop-out | Routing still applies; inspect account-scoped behavior because item identity can change |
| Several open draft windows | Each uses its own original-account baseline |
| Cc/Bcc mapped differently | Send canceled, draft retained; no delivery |
| Mixed mapped/unmapped recipients | Send canceled; whole-message ambiguity never picks a sender silently |
| Unresolved recipient or group | No automatic guess; Send canceled until resolved/expanded into supported entries |
| Change recipients/remove all matching recipients | Original sender restored while the draft remains in the same open session |
| Manually change From on a matched draft | Configured sender restored by next evaluation or Send |
| Send immediately after changing recipients | Final synchronous evaluation applies the current mapping |
| Configured profile account missing | Send canceled with explanation; no silent fallback |
| Delegated sender unavailable in GAL | Send canceled before submission |
| Valid delegated mailbox, Send As granted | Received message shows shared mailbox as sender; inspect headers |
| Only Send on Behalf granted | Received message identifies delegate on behalf of mailbox; inspect headers |
| Neither delegated sending right | Exchange rejection/NDR expected even if local property setting succeeds |
| Both sending rights | Verify expected Send As behavior and actual Sent Items location |
| Save, close, reopen mapped draft | Verify new baseline and account-scoped rules; original baseline does not persist across closure |
| Signature, attachments, labels/encryption | Inspect that native compose content and intended sender/signature remain correct; protected/custom forms are not certified |
| Offline/unavailable directory | Fail safely if sender/recipient cannot resolve; queued transport behavior needs profile-specific validation |
| Malformed/missing rules file at startup | Startup explanation and blocked mail sending; Settings permits recovery and saves a backup of the previous file |
| External file edit | No change until Outlook restart; settings edited through the ribbon apply immediately |
| Disable COM add-in directly/uninstall | Future automatic changes stop; existing saved drafts retain their current From |
| Restart and update | Add-in reloads, user rules are preserved |

Verify both 32-bit and 64-bit installations if you plan to deploy to both. Check Outlook performance/disabled-add-in reports during a pilot with your normal number of open drafts. Large directories, slow resolution, and many open drafts can add UI-thread work. Observe behavior alongside other sender/signature/security add-ins; callback order is not a transport-level guarantee.

Installer/uninstaller scripts are syntax-checked; real installation and removal are not performed by the automated suite. Detection tests cover x86/x64 executable headers and invalid/unsupported formats. Verify automatic detection and removal in both Outlook bitnesses on disposable Windows profiles before release.

Version 1.2.6 defers the initial sender write to a 50 ms Windows Forms timer tick after NewInspector, then returns to 250 ms scanning. It avoids the pre-display write introduced in 1.2.5, which can leave visible From stale while the underlying alias is correct. This addresses the suspected timing cause; live UI refresh after repeated discard/reopen cycles still needs verification. Timer intervals are not guaranteed wall-clock latency.
