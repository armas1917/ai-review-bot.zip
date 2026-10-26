"""
AI Review Bot - automates task allocation, review scheduling and feedback emails.

STEP 1  python review_bot.py assign            -> task + deadline + review slot email for every student
STEP 2  python review_bot.py feedback          -> after review: mistakes + feedback + next review slot email

Default is DRY-RUN (emails saved in ./outbox as .txt files). Add --send to really email.

Excel input (data/Student_Task_List.xlsx):
  Sheet "Students": Student Register Number | Name | Email | Task
  Sheet "Reviews" (fill after the meeting, used by STEP 2):
        Student Register Number | Submission | Explanation | Reviewer Notes (optional)

Env vars:
  ANTHROPIC_API_KEY  -> AI writes the emails/feedback (without it, built-in templates are used)
  SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, MAIL_FROM   -> needed only with --send
  MEET_LINK          -> your Google Meet / Zoom link (one link, or edit make_link())
"""
import argparse, os, smtplib, ssl, sys
from datetime import datetime, timedelta, time
from email.message import EmailMessage
from pathlib import Path

import pandas as pd

EXCEL = "data/Student_Task_List.xlsx"
OUTBOX = Path("outbox")
SCHEDULE_FILE = "review_schedule.xlsx"
MODEL = "claude-sonnet-5-5"

SLOT_MINUTES = 10                       # one student per slot
DAY_WINDOWS = [(time(10, 0), time(13, 0)), (time(14, 0), time(17, 0))]
MEET_LINK = os.getenv("MEET_LINK", "https://meet.google.com/your-link-here")

TASK_HELP = {  # fallback instructions if no AI key
    "Poster design": "Design an A4/Instagram poster with event name, date, time, venue and a clear tagline.",
    "Video editing": "Edit a 15-20 second promo video with title, music and event details.",
    "Content writing": "Write the event description, tagline and 3 social media captions.",
    "Registration desk": "Prepare the registration sheet, attendee list format and desk flow.",
    "Stage decoration": "Plan the stage layout, backdrop and decoration list with a sketch.",
    "Photography": "Prepare a shot list and cover sample photos for the event.",
    "Social media posts": "Create 3 post designs and a posting calendar.",
    "Certificate preparation": "Design the certificate template and prepare names from the list.",
    "Guest coordination": "Prepare the guest list, welcome plan and contact schedule.",
    "Announcements / emceeing": "Write the emcee script with announcements and a timeline.",
}

# ----------------------------------------------------------------- AI helper
def ask_ai(prompt: str, fallback: str) -> str:
    key = os.getenv("ANTHROPIC_API_KEY")
    if not key:
        return fallback
    try:
        import anthropic
        r = anthropic.Anthropic(api_key=key).messages.create(
            model=MODEL, max_tokens=700,
            system=("You are a friendly review assistant for college students. "
                    "Explain in very simple English, like a helpful mentor. Keep it short."),
            messages=[{"role": "user", "content": prompt}],
        )
        return r.content[0].text.strip()
    except Exception as e:
        print("  [AI failed, using template]", e)
        return fallback

# ----------------------------------------------------------------- scheduling
def slot_generator(start_day: datetime):
    """Yield consecutive review slots, skipping Sundays."""
    day = start_day.replace(hour=0, minute=0, second=0, microsecond=0)
    while True:
        if day.weekday() != 6:
            for s, e in DAY_WINDOWS:
                t = datetime.combine(day.date(), s)
                end = datetime.combine(day.date(), e)
                while t + timedelta(minutes=SLOT_MINUTES) <= end:
                    yield t
                    t += timedelta(minutes=SLOT_MINUTES)
        day += timedelta(days=1)

def fmt(dt: datetime) -> str:
    return dt.strftime("%A, %d %B %Y at %I:%M %p")

def make_link(reg) -> str:
    return MEET_LINK  # replace with a unique link per student if you generate them

# ----------------------------------------------------------------- email
def deliver(to, subject, body, send):
    if send:
        msg = EmailMessage()
        msg["From"] = os.environ["MAIL_FROM"]
        msg["To"], msg["Subject"] = to, subject
        msg.set_content(body)
        with smtplib.SMTP_SSL(os.environ["SMTP_HOST"], int(os.getenv("SMTP_PORT", 465)),
                              context=ssl.create_default_context()) as s:
            s.login(os.environ["SMTP_USER"], os.environ["SMTP_PASS"])
            s.send_message(msg)
        print(f"  sent -> {to}")
    else:
        OUTBOX.mkdir(exist_ok=True)
        fname = OUTBOX / f"{to.replace('@', '_at_')}_{abs(hash(subject)) % 10000}.txt"
        fname.write_text(f"To: {to}\nSubject: {subject}\n\n{body}\n", encoding="utf-8")
        print(f"  saved -> {fname}")

def load_students():
    df = pd.read_excel(EXCEL, sheet_name="Students", dtype=str)
    df.columns = [c.strip() for c in df.columns]
    return df

# ----------------------------------------------------------------- STEP 1
def assign(args):
    df = load_students()
    deadline = datetime.now() + timedelta(days=args.deadline_days)
    deadline = deadline.replace(hour=17, minute=0, second=0, microsecond=0)
    slots = slot_generator(deadline + timedelta(days=1))   # reviews start the day after the deadline
    rows = []
    for _, st in df.iterrows():
        name, email, task, reg = st["Name"], st["Email"], st["Task"], st["Student Register Number"]
        slot = next(slots)
        link = make_link(reg)
        help_txt = TASK_HELP.get(task, "Complete your task as discussed.")
        explain = ask_ai(
            f"Student {name} got the task '{task}' for a college event. Explain in 4-5 simple lines "
            f"what to do, what a good output looks like, and what they must be ready to explain in the review.",
            help_txt)
        body = (f"Hi {name},\n\nYour task: {task}\n\n{explain}\n\n"
                f"DEADLINE: {fmt(deadline)}\n"
                f"REVIEW MEETING: {fmt(slot)} ({SLOT_MINUTES} min)\nMeeting link: {link}\n\n"
                f"In the review, explain your work in a very simple way: what you made, why, and what you learned.\n"
                f"You will get feedback by email after the review.\n\nAll the best!\nReview Bot")
        deliver(email, f"Your task: {task} | Deadline & review time", body, args.send)
        rows.append(dict(Register=reg, Name=name, Email=email, Task=task,
                         Deadline=deadline, Review=slot, Link=link))
    pd.DataFrame(rows).to_excel(SCHEDULE_FILE, index=False)
    print(f"\nDone. {len(rows)} students. Schedule saved to {SCHEDULE_FILE}")

# ----------------------------------------------------------------- STEP 2
def feedback(args):
    students = load_students().set_index("Student Register Number")
    reviews = pd.read_excel(EXCEL, sheet_name="Reviews", dtype=str).fillna("")
    reviews.columns = [c.strip() for c in reviews.columns]
    sched = pd.read_excel(SCHEDULE_FILE)
    last = pd.to_datetime(sched["Review"]).max()
    slots = slot_generator(last + timedelta(days=args.gap_days))
    for _, r in reviews.iterrows():
        reg = r["Student Register Number"]
        if reg not in students.index:
            print("  unknown register no:", reg); continue
        st = students.loc[reg]
        name, task = st["Name"], st["Task"]
        slot = next(slots)
        fb = ask_ai(
            f"Task: {task}\nStudent: {name}\nSubmission (what they delivered): {r.get('Submission','')}\n"
            f"Student's explanation in the review: {r.get('Explanation','')}\n"
            f"Reviewer notes: {r.get('Reviewer Notes','')}\n\n"
            "Write feedback in simple English with 3 parts: 1) What was good, 2) Mistakes / what is missing "
            "(be specific), 3) What to fix before the next review. Be kind but honest.",
            f"What was good: you completed the task.\nMistakes: please check "
            f"{r.get('Reviewer Notes','') or 'quality, clarity and missing details'}.\nFix these before the next review.")
        body = (f"Hi {name},\n\nThanks for attending the review of '{task}'.\n\nFEEDBACK\n{fb}\n\n"
                f"NEXT REVIEW: {fmt(slot)}\nMeeting link: {make_link(reg)}\n\nKeep going!\nReview Bot")
        deliver(st["Email"], f"Feedback on your task: {task} + next review time", body, args.send)

# ----------------------------------------------------------------- main
if __name__ == "__main__":
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("assign"); a.add_argument("--deadline-days", type=int, default=3)
    a.add_argument("--send", action="store_true"); a.set_defaults(fn=assign)
    f = sub.add_parser("feedback"); f.add_argument("--gap-days", type=int, default=3)
    f.add_argument("--send", action="store_true"); f.set_defaults(fn=feedback)
    args = p.parse_args()
    if not Path(EXCEL).exists():
        sys.exit(f"Put {EXCEL} in this folder first.")
    args.fn(args)
