# Automate-Bulk-Email
Automate you Email Campaign with out Subscription 
# Google Sheets Email Campaign & Bounce Tracker

A beginner-friendly Google Apps Script solution that sends customized HTML emails directly from a Google Sheet, tracks daily send limits, attaches Google Drive files, and automatically detects bounced emails from Gmail[cite: 1, 2].

## Features

- **No-Code Customization:** Edit the subject line, Google Drive attachment ID, and email body easily without programming experience[cite: 1, 2].
- **Full-Name Greetings:** Uses the value in the Full Name column for greetings (falling back to `"there"` if blank)[cite: 2].
- **Rich HTML Formatting:** Supports `<p>` paragraphs, `<b>` bold text, `<ul><li>` bullet points, and dynamic unsubscribe links with the recipient's email[cite: 1, 2].
- **Daily Limit Tracking:** Checks the active user account and tracks daily sent counts using script properties[cite: 2].
- **Bounce Tracking:** Scans Gmail for recent delivery failure notifications (`mailer-daemon`), updates matching `SENT` rows in Column C to `BOUNCED`, highlights the row light red (`#f4c7c7`), and marks the bounce thread as read[cite: 1, 2].

## Spreadsheet Setup

Create a Google Sheet with data starting on **Row 2**:

| Column A | Column B | Column C |
| :--- | :--- | :--- |
| Email Address | Full Name | Status (blank or `PENDING` to send; updates to `SENT`, `FAILED`, or `BOUNCED`)[cite: 1, 2] |

## How to Install and Use

1. Open your Google Sheet and click **Extensions** > **Apps Script**.
2. Delete any default code in the editor, paste the contents of `Code.gs`, and save the project.
3. Edit `emailSubject`, `googleDriveFileId` (paste the file ID from the Google Drive sharing URL, or leave `""` for no attachment), and the `htmlBody` block in `sendCustomEmails()`[cite: 1, 2]. Replace `https://yourwebsite.com/unsubscribe` with your real unsubscribe URL[cite: 1, 2].
4. Select **`sendCustomEmails`** in the Apps Script toolbar and click **Run**. Approve the Google permissions prompts the first time you run it[cite: 1].
5. Select **`checkBouncedEmails`** in the toolbar and click **Run** to scan recent bounce notifications and update matching rows[cite: 1, 2].
