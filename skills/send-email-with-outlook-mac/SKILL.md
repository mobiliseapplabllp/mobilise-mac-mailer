---
name: send-email-with-outlook-mac
description: Send or draft email through Microsoft Outlook on the user's Mac using the Mobilise Mac Mailer tools. Use when the user asks to email someone, send a follow-up or reply-style message, or prepare a draft in Outlook.
---

# Sending email with Outlook on a Mac

The Mobilise Mac Mailer tools drive the Microsoft Outlook app on the user's Mac. Mail goes out from Outlook's default account and Outlook adds the user's signature.

## Before sending

1. Write the email in the conversation first: recipients (to, cc, bcc), subject, body, and any attachments.
2. Use only addresses the user gave you or that come from their own material for this task. Never guess an address.
3. Call `send_email` only after the user approves that exact message. If they haven't approved the final wording, or they want to look it over in Outlook, call `create_draft` instead.
4. Don't add a sign-off block with contact details: Outlook appends the user's signature. A short closing such as "Regards," followed by the user's name is fine.

## Writing the call

- Use `body_format: "text"` for normal emails. Blank lines become paragraphs.
- Use `body_format: "html"` only when the user needs formatting such as lists or bold text.
- Attachments must be absolute paths to files on the Mac.

## When something goes wrong

- **Blocked by the allowlist:** tell the user which address was blocked and that they can add it to the allowlist file named in the error. Don't try to send another way.
- **macOS blocked access to Outlook:** ask the user to allow control of Microsoft Outlook in System Settings > Privacy & Security > Automation, then call `check_setup`.
- **Hourly limit reached:** tell the user, and offer `create_draft` instead.

## After sending

Tell the user what went out: subject, recipients and attachment count. Use `recent_activity` if they ask what has been sent. Emails sent this way may not appear in Outlook's Sent Items; the activity log is the record.
