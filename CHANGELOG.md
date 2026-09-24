# Changes

## 1.2.6 - defer first write until UI-loop retry

- Replace synchronous pre-display sender writes with a 50 ms scheduled scan, then resume 250 ms fallback.
- Keep alias verification unchanged.
- Add discard/reopen lifecycle coverage and verify no sender write occurs inside NewInspector. 113 tests pass; live Outlook UI behavior remains to be confirmed.

## 1.2.5 - faster compose selection

- Apply the sender from Inspectors.NewInspector before the compose window appears when the draft is ready.
- Reduce fallback scanning from 1000 ms to 250 ms, including inline replies and initialization retries.
- Retain and detach the event source/delegate; failed optional event hookup preserves timer and send checks and is reported by Check draft.
- Add callback regression tests; 106 tests pass. Actual visual timing remains dependent on Outlook and requires live verification.

## 1.2.4 - observed Outlook alias normalization

- Accept the reported pair of empty address-type/email-address strings only when literal alias SMTP, full one-off entry ID, search key, and account remain exact.
- Keep rejecting missing, conflicting, or partially cleared address metadata; preserve strict snapshot restoration.
- Avoid repeated saves for this normalized identity. Add regression and rejection tests; 99 add-in tests pass.
- Visible From refresh and delivered alias still require live verification.

## 1.2.3 - alias verification correction

- Stop treating represented display name as an SMTP address. Continue requiring the selected account and all five alias address properties to match exactly.
- Include expected/read-back identity differences captured before rollback in Check draft errors.
- Add tests for display-name normalization, unchanged-draft behavior, rejected SMTP rewrite with diagnostic evidence, and missing identity metadata.
- Live profile compatibility remains unverified; no tenant settings or mail delivery were changed.

## 1.2.2 - draft diagnostics

- Add a read-only Check draft ribbon action showing selection, current sender properties, and automatic scan/selection errors with HRESULTs.
- Keep draft errors in memory while the draft is open; no recipient/body logging or message submission.
- This diagnostic update does not claim to fix live alias selection failures; use its report to identify the failing operation.

## 1.2.1 - alias validation and installation fixes

- Verify account and alias membership in one proxy list without requiring the account to equal the directory primary address.
- Improve missing sender-label errors.
- Detect classic Outlook executable bitness automatically; retain explicit override.
- Add double-click install/uninstall launchers and Installed Apps registration.
- Verify uninstall across both registry views, report machine-wide remnants, preserve settings by default, and refuse removal while Outlook is running.

## 1.2.0 - alias pilot

- Separate aliases from delegated/shared mailboxes in configuration and the Senders editor.
- Add Use as default per owning Outlook account, with no recipient rule required and selection on empty drafts.
- Verify alias membership using the selected account's address-book proxy addresses.
- Write a literal SMTP represented-sender identity through documented MAPI properties, with a Unicode one-off entry ID; do not resolve the alias into the primary mailbox address.
- Save and verify sender metadata, restore the previous identity on failure, and reject observable client-side rewriting.
- Preserve exact represented-sender metadata when restoring drafts or switching back to the primary account.
- Read version 1 configurations; save version 2. Do not silently reclassify old delegated entries as aliases.
- Add ownership, default routing, persistence, rollback and rewrite tests. Live Exchange delivery is still unverified.

## 1.1.0

- Add AutoFrom > Settings in classic Outlook's main and mail compose ribbons.
- Add a graphical editor for allowed senders, Outlook accounts, shared mailboxes, and recipient/domain/account rules.
- Preview unsaved rule selection without sending a message.
- Validate and atomically save settings with a backup; apply changes without restarting Outlook.
- Restore managed open drafts before applying a different policy or disabling selection.
- Recover invalid settings from inside Outlook.
- Add configuration persistence, ribbon, and settings editor tests.

## 1.0.0

- Initial automatic sender selection for classic Outlook using the Outlook Object Model.
- Recipient/domain/original-account matching, delegated sender routing and send-time checks.
