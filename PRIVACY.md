# Privacy policy for Mobilise Mac Mailer

Last updated: 2 October 2026

Mobilise Mac Mailer is published by Mobilise App Lab Limited. It runs entirely on your Mac.

## What the plugin handles

- **Email content you ask Claude to send or draft**: recipients, subject, body and attachment paths. The plugin passes these to the Microsoft Outlook app on your Mac, which sends the email from your own account. The plugin does not keep a copy of the body.
- **Activity log**: for each send or draft, the plugin records the time, the action, the To and CC addresses, the number of BCC recipients, the subject and the number of attachments in `~/.mobilise-mac-mailer/activity-log.tsv` on your Mac, so you can check what was sent.
- **Allowlist**: if you create `~/.mobilise-mac-mailer/allowed-recipients.txt`, the plugin reads it to decide which recipients are allowed.

## What the plugin does not do

- It makes no network connections of its own, and sends no data to Mobilise App Lab Limited or anyone else.
- It does not collect analytics or telemetry.
- It does not read your mailbox, contacts, calendar, Claude conversations, chat history or files other than the attachments you name.

## Retention and deletion

The activity log and allowlist stay on your Mac until you delete them. Uninstalling the plugin does not delete them; remove the `~/.mobilise-mac-mailer` folder to delete them. Emails you send are handled by Microsoft Outlook and your email provider under their own terms.

## Contact

Questions about this policy: open an issue at https://github.com/mobiliseapplabllp/mobilise-mac-mailer/issues.
