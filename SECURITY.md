# Security and data privacy

MarbleDeliveryPro is a local desktop program. It may hold personally identifiable customer and employee data.

- Never commit `*.db`, backups, attachments, generated reports, exports, logs, API keys, or tokens.
- The default built-in administrator does not have a configured password; passwordless accounts are blocked from authentication in 10.0.5. Configure an administrator password before requiring login.
- Login is not a substitute for Windows filesystem security or full-disk encryption.
- Protect `%LOCALAPPDATA%\MarbleDeliveryPro` with your Windows account and encrypted backups.
- Third-party API integrations are opt-in; protect the configured API tokens. This release does not claim encrypted secrets at rest.
- Confirm the vendor and SHA256 manifest before installing. Unsigned EXE installers may trigger SmartScreen until they are signed and gain reputation.
- To report security problems, contact the maintainer privately through GitHub rather than publishing user data in a public issue.
