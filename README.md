# Mobilise Mac Mailer

Mobilise Mac Mailer lets Claude send and draft email through the Microsoft Outlook app on your Mac. Messages go out from your own Outlook account with your normal signature, so there is no mailbox password to share, no cloud relay and no third-party service. It works with both the new and the legacy Outlook for Mac.

## What it can do

| Tool | What it does |
| --- | --- |
| `send_email` | Sends an email now, with optional CC, BCC and attachments |
| `create_draft` | Opens a filled-in email in Outlook for you to review and send yourself |
| `check_setup` | Confirms Outlook is installed and macOS allows control; sends nothing |
| `list_allowed_recipients` | Shows your recipient allowlist, if you set one |
| `recent_activity` | Lists recent sends and drafts from the local activity log |

The plugin also includes a skill that tells Claude to show you each email and get your approval before it sends anything.

## Requirements

- macOS with Microsoft Outlook installed and signed in
- Node.js 18 or later available as `node`
- A Claude surface that runs local plugin servers on your Mac: Cowork in the Claude desktop app, or Claude Code. Chat on claude.ai does not run local servers.

The first time Claude sends or drafts an email, macOS asks whether the app may control Microsoft Outlook. Choose **OK**. You can change this later in System Settings > Privacy & Security > Automation.

## Safety controls

- **Recipient allowlist (optional).** Create `~/.mobilise-mac-mailer/allowed-recipients.txt` with one address per line, or `@yourcompany.com` to allow a whole domain. When the file exists, Claude can only email those recipients. Delete the file to remove the restriction.
- **Hourly limit.** At most 20 emails per hour by default. Set the `MAC_MAILER_MAX_SENDS_PER_HOUR` environment variable to change it.
- **Activity log.** Every send and draft is recorded in `~/.mobilise-mac-mailer/activity-log.tsv`. Emails sent this way may not appear in Outlook's Sent Items, so this log is the record.
- **Data folder.** Set `MAC_MAILER_DATA_DIR` to keep the allowlist and log somewhere else.

## What it runs and sends

- It runs `server/server.js` with Node.js and `server/outlook.applescript` with `/usr/bin/osascript`.
- It makes no network connections of its own and sends no telemetry.
- Email leaves your Mac only through Outlook, to the recipients in the tool call.
- It writes temporary files to your system temp folder while handing a message to Outlook, deletes them right after, and writes the activity log described above.

See [PRIVACY.md](PRIVACY.md) for the privacy policy.

## Troubleshooting

- **"macOS blocked access to Outlook"**: allow control of Microsoft Outlook in System Settings > Privacy & Security > Automation, then ask Claude to run `check_setup`.
- **"Outlook isn't installed or couldn't be found"**: install Microsoft Outlook and sign in.
- **An email went to the wrong account**: Outlook sends from its default account. Change the default in Outlook's settings.

## License

MIT. See [LICENSE](LICENSE).
