# AI Review Bot

Automates the student task-review process for events: assigns each student a task, emails the deadline and a review time slot, and after the review sends feedback (mistakes + what to fix) with the next review time.

## What is inside

| Folder / file | What it does |
|---|---|
| `web/review_bot.html` | Chatbot page: type `assign tasks`, `email student 5`, `feedback student 5`. Open the file in a browser. |
| `web/review_dashboard.html` | Dashboard: progress cards, charts, student table with status. Open in a browser. |
| `apps_script/Code.gs` | Google Sheets menu **Review Bot** that sends the emails from your own Gmail. |
| `review_bot.py` | Python version: sends emails by SMTP, optional AI-written feedback (Claude API). |
| `data/Student_Task_List.xlsx` | Sample data (fake `@example.com` addresses). Sheets: `Students`, `Reviews`. |

## Excel format

Sheet `Students`: `Student Register Number | Name | Email | Task`
Sheet `Reviews` (filled after the review meeting): `Student Register Number | Submission | Explanation | Reviewer Notes`

## Option 1: Google Sheets (easiest, sends real emails)

1. Upload `data/Student_Task_List.xlsx` to Google Drive and open with Google Sheets. Put real student emails in column C.
2. `Extensions -> Apps Script`, paste `apps_script/Code.gs`, set `MEET_LINK`, save.
3. Refresh the sheet. Use the **Review Bot** menu: test email, send task emails, send feedback emails.
4. Emails to `@example.com` addresses are skipped. Each student is never emailed twice (see the Status column).
5. Free Gmail limit: about 100 emails per day.

## Option 2: Python

```bash
pip install -r requirements.txt
python review_bot.py assign              # dry run: emails saved in ./outbox
python review_bot.py assign --send       # really send
python review_bot.py feedback --send     # after filling the Reviews sheet
```

Environment variables for sending: `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS` (Gmail App Password), `MAIL_FROM`, `MEET_LINK`.
Optional: `ANTHROPIC_API_KEY` for AI-written explanations and feedback.

## Option 3: Web pages

Just open `web/review_bot.html` and `web/review_dashboard.html` in a browser. They show and prepare emails but cannot send mail themselves (browsers cannot send email). You can upload your own Excel file with the paperclip / Upload button.

## Privacy

Never commit real student names, emails or register numbers to a public repository. Keep your real Excel file local (it is in `.gitignore`) and never commit passwords, App Passwords or API keys.
