-- Outlook bridge for Mobilise Mac Mailer (called by server.js).
--
-- Arguments: <mode> <subject> <body_file> <recipients_file> <attachments_file>
--   mode              send  - send the message now
--                     draft - open the message in Outlook without sending
--                     build - build the message and report its contents, without
--                             opening or sending it (used for testing)
--   body_file         UTF-8 HTML body
--   recipients_file   one line per recipient: to:<address>, cc:<address> or bcc:<address>
--   attachments_file  one absolute file path per line (may be empty)
--
-- Outlook sends from its default account and adds the user's signature.

on run argv
	if (count of argv) < 5 then error "Expected 5 arguments: mode, subject, body file, recipients file, attachments file"
	set theMode to (item 1 of argv) as text
	set theSubject to (item 2 of argv) as text
	set bodyPath to (item 3 of argv) as text
	set recipientsPath to (item 4 of argv) as text
	set attachmentsPath to (item 5 of argv) as text
	if theMode is not in {"send", "draft", "build"} then error "Unknown mode: " & theMode

	set theBody to do shell script ("cat " & quoted form of bodyPath) without altering line endings
	set recipientLines to paragraphs of (do shell script "cat " & quoted form of recipientsPath)
	set attachmentLines to paragraphs of (do shell script "cat " & quoted form of attachmentsPath)

	tell application "Microsoft Outlook"
		set newMessage to make new outgoing message with properties {subject:theSubject, content:theBody}
		repeat with rawLine in recipientLines
			set lineText to contents of rawLine
			if lineText starts with "to:" then
				make new to recipient at newMessage with properties {email address:{address:(text 4 thru -1 of lineText)}}
			else if lineText starts with "cc:" then
				make new cc recipient at newMessage with properties {email address:{address:(text 4 thru -1 of lineText)}}
			else if lineText starts with "bcc:" then
				make new bcc recipient at newMessage with properties {email address:{address:(text 5 thru -1 of lineText)}}
			end if
		end repeat
		repeat with rawPath in attachmentLines
			set p to contents of rawPath
			if p is not "" then make new attachment at newMessage with properties {file:(POSIX file p)}
		end repeat

		if theMode is "send" then
			send newMessage
			return "sent"
		else if theMode is "draft" then
			open newMessage
			activate
			return "draft opened"
		else
			return "built: " & (count of to recipients of newMessage) & " to, " & (count of cc recipients of newMessage) & " cc, " & (count of bcc recipients of newMessage) & " bcc, " & (count of attachments of newMessage) & " attachments"
		end if
	end tell
end run
