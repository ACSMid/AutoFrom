# Preparing a GitHub release

PowerShell remains useful for developers and pilot testing. For general users, the recommended release experience is a signed **Setup.exe** with a normal installation wizard and an uninstall entry in Windows Settings. An MSI can be a separate option if an organization needs centrally managed deployment. The current archive contains the development installer scripts; a signed GUI installer has **not** been built yet.

An Inno Setup installer is a practical candidate for this per-user COM add-in. It should install with user privileges, register the CLR COM class for Outlook's actual bitness, preserve existing settings, and remove only its own registration on uninstall. [Inno Setup registry/bitness documentation](https://jrsoftware.org/ishelp/topic_32vs64bitinstalls.htm). Do not infer Outlook bitness solely from Windows bitness.

Before publishing an installable release:

1. Complete the live Outlook tests in TESTING.md, including ribbon visibility and Settings callbacks on 32-bit and 64-bit Outlook. Test upgrading from the previous assembly version and recovering malformed settings.
2. Build a setup package that detects classic Outlook, its bitness, .NET Framework 4.8, and whether Outlook is running. If detection is ambiguous, offer an explicit choice. Require Outlook to close before replacing the DLL; do not terminate it and risk losing drafts.
3. Preserve `%LOCALAPPDATA%\AutoFrom\rules.json` and its backup through upgrades. Offer a clear settings-retention choice on uninstall. Do not grant mailbox permissions or silently change security policy.
4. Sign the DLL and installer with the publisher's Authenticode identity and timestamp the signatures. Signing credentials belong in a protected signing service or release secrets, never in the repository. Signing identifies the publisher; it does not eliminate all SmartScreen or enterprise-policy restrictions.
5. Include LICENSE.md (PolyForm Noncommercial 1.0.0) in source and binary distributions. Describe the project as source available, not open source. The pilot binaries do not claim a verified publisher.
6. Commit source, tests, documentation and inert/example configuration to GitHub. `.gitignore` excludes generated binaries, developer configuration copies, backups and signing keys. Do not commit your installed configuration or real mailbox/domain lists. Put installer binaries and checksums on a tagged GitHub Release rather than committing them to source history.

No credentials, private mailbox configuration, or signing certificate are included in the distribution. Public repository and commit metadata identify the publishing GitHub account.

## Ribbon implementation

The AutoFrom tab is supplied for `Microsoft.Outlook.Explorer` and `Microsoft.Outlook.Mail.Compose`. `Connect` implements both lifecycle and ribbon interfaces, with a COM-visible dispatch interface for its Settings callback. The settings editor offers senders, rules and a local test page. [Microsoft's Outlook ribbon interface documentation](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/office-fluent-ui-extensibility/implementing-the-iribbonextensibility-interface).
