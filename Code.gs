const MEET_LINK = "https://meet.google.com/your-link-here";
const DEADLINE_DAYS = 3, GAP_DAYS = 3, SLOT_MIN = 10;

const HELP = {
  "Poster design": "Design a poster with event name, date, time, venue and a clear tagline.",
  "Video editing": "Edit a 15-20 second promo video with title, music and event details.",
  "Content writing": "Write the event description, a tagline and 3 social media captions.",
  "Registration desk": "Prepare the registration sheet, attendee list and desk flow.",
  "Stage decoration": "Plan the stage layout, backdrop and decoration list with a sketch.",
  "Photography": "Make a shot list and capture sample photos of the event.",
  "Social media posts": "Create 3 post designs and a posting calendar.",
  "Certificate preparation": "Design a certificate template and fill the names from the list.",
  "Guest coordination": "Prepare the guest list, welcome plan and contact schedule.",
  "Announcements / emceeing": "Write the emcee script with announcements and a timeline."
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Review Bot")
    .addItem("1. Send me a test email", "sendTestEmail")
    .addItem("2. Send task emails", "sendTaskEmails")
    .addItem("3. Send feedback emails", "sendFeedbackEmails")
    .addToUi();
}

function fmt(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "EEEE, dd MMMM yyyy 'at' hh:mm a");
}

// Review slots: 10-1 and 2-5, 10 minutes each, no Sundays, never before 'start'
function makeSlots(start, n) {
  const out = []; let day = new Date(start); day.setHours(0, 0, 0, 0);
  while (out.length < n) {
    if (day.getDay() !== 0) {
      [[10, 13], [14, 17]].forEach(w => {
        let t = new Date(day); t.setHours(w[0], 0, 0, 0);
        const e = new Date(day); e.setHours(w[1], 0, 0, 0);
        while (t.getTime() + SLOT_MIN * 60000 <= e.getTime() && out.length < n) {
          if (t.getTime() >= start.getTime()) out.push(new Date(t));
          t = new Date(t.getTime() + SLOT_MIN * 60000);
        }
      });
    }
    day.setDate(day.getDate() + 1);
  }
  return out;
}

function validEmail(e) {
  e = String(e).trim();
  return e && e.indexOf("@") > 0 && e.indexOf("@example.com") === -1;
}

function taskBody(name, task, deadline, slot) {
  return "Hi " + name + ",\n\nYour task: " + task + "\n\nWhat to do: " + (HELP[task] || "Complete your task as discussed.") +
    "\n\nDEADLINE: " + fmt(deadline) + "\nREVIEW MEETING: " + fmt(slot) + " (" + SLOT_MIN + " min)\nMeeting link: " + MEET_LINK +
    "\n\nIn the review, explain your work in very simple words: what you made, why, and what you learned. You will get feedback by email after the review.\n\nAll the best!\nReview Bot";
}

function sendTestEmail() {
  const me = Session.getActiveUser().getEmail();
  const dl = new Date(Date.now() + DEADLINE_DAYS * 864e5); dl.setHours(17, 0, 0, 0);
  const slot = makeSlots(new Date(dl.getTime() + 864e5), 1)[0];
  MailApp.sendEmail(me, "TEST: Your task: Poster design | Deadline & review time", taskBody("Test Student", "Poster design", dl, slot));
  SpreadsheetApp.getUi().alert("Test email sent to " + me + ". Check your inbox (and Spam).");
}

function sendTaskEmails() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName("Students");
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  const status = sh.getRange(2, 7, rows.length, 1).getValues();
  const todo = rows.map((r, i) => i).filter(i => validEmail(rows[i][2]) && status[i][0] !== "Task email sent");
  const skipped = rows.length - todo.length;
  if (!todo.length) { ui.alert("Nothing to send. All valid students already emailed, or emails are missing."); return; }
  if (ui.alert("Send task emails to " + todo.length + " students?\n(" + skipped + " skipped: fake/missing email or already sent)", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const deadline = new Date(Date.now() + DEADLINE_DAYS * 864e5); deadline.setHours(17, 0, 0, 0);
  const slots = makeSlots(new Date(deadline.getTime() + 864e5), rows.length);
  sh.getRange(1, 5, 1, 3).setValues([["Deadline", "Review time", "Status"]]);
  let sent = 0;
  todo.forEach(i => {
    const r = rows[i];
    MailApp.sendEmail(String(r[2]).trim(), "Your task: " + r[3] + " | Deadline & review time", taskBody(r[1], r[3], deadline, slots[i]));
    sh.getRange(i + 2, 5, 1, 3).setValues([[fmt(deadline), fmt(slots[i]), "Task email sent"]]);
    sent++;
  });
  ui.alert("Done. " + sent + " task emails sent. Check the Status column.");
}

function sendFeedbackEmails() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  let rv = ss.getSheetByName("Reviews");
  if (!rv) {
    rv = ss.insertSheet("Reviews");
    rv.appendRow(["Student Register Number", "Submission", "Explanation", "Reviewer Notes (mistakes, comma separated)", "Status", "Next review"]);
    ui.alert("I created a Reviews sheet. Fill it after the review meetings, then run this again.");
    return;
  }
  const info = {};
  ss.getSheetByName("Students").getDataRange().getValues().slice(1).forEach(r => info[String(r[0]).trim()] = { name: r[1], email: r[2], task: r[3] });
  const n = rv.getLastRow() - 1;
  if (n < 1) { ui.alert("Reviews sheet is empty."); return; }
  const rows = rv.getRange(2, 1, n, 6).getValues();
  const todo = rows.map((r, i) => i).filter(i => {
    const s = info[String(rows[i][0]).trim()];
    return s && validEmail(s.email) && String(rows[i][3]).trim() && rows[i][4] !== "Feedback sent";
  });
  if (!todo.length) { ui.alert("Nothing to send. Check register numbers, emails and the mistakes column."); return; }
  if (ui.alert("Send feedback emails to " + todo.length + " students?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  let start = new Date(Date.now() + GAP_DAYS * 864e5);
  rows.forEach(r => { if (r[5] instanceof Date && r[5].getTime() + SLOT_MIN * 60000 > start.getTime()) start = new Date(r[5].getTime() + SLOT_MIN * 60000); });
  const slots = makeSlots(start, todo.length);
  rv.getRange(1, 5, 1, 2).setValues([["Status", "Next review"]]);
  todo.forEach((i, k) => {
    const r = rows[i], s = info[String(r[0]).trim()];
    const mistakes = String(r[3]).split(/[,;\n]+/).map(x => x.trim()).filter(x => x).map(x => "- " + x).join("\n");
    const body = "Hi " + s.name + ",\n\nThanks for attending the review of \"" + s.task + "\" and explaining your work.\n\n" +
      "WHAT WENT WELL\n- You completed the task and presented it.\n\nMISTAKES / WHAT IS MISSING\n" + mistakes +
      "\n\nWHAT TO DO NEXT\nPlease fix the points above and be ready to explain the changes simply.\n\n" +
      "NEXT REVIEW: " + fmt(slots[k]) + "\nMeeting link: " + MEET_LINK + "\n\nKeep going!\nReview Bot";
    MailApp.sendEmail(String(s.email).trim(), "Feedback on your task: " + s.task + " + next review time", body);
    rv.getRange(i + 2, 5, 1, 2).setValues([["Feedback sent", slots[k]]]);
    rv.getRange(i + 2, 6).setNumberFormat("dd MMM yyyy, hh:mm a");
  });
  ui.alert("Done. " + todo.length + " feedback emails sent.");
}
