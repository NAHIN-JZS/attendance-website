# Class attendance → Google Sheet

A one-page site for marking attendance on your phone. It writes straight into a
Google Sheet you own. No database, no server, no paid anything.

```
index.html            the site (this is what you deploy)
apps-script/Code.gs    the bit that lives inside your Google Sheet
```

---

## 1. Make the sheet

1. Create a new Google Sheet. Name it whatever you like.
2. **Extensions → Apps Script.** Delete the sample code, paste in everything from
   `apps-script/Code.gs`, and save.
3. **Project Settings** (gear icon on the left) → set the time zone to
   **Asia/Dhaka**. This is what dates and timestamps use.
4. Back in the editor, pick `setupSheets` from the function dropdown and click
   **Run**. Approve the permission prompt the first time — it only ever touches
   this one sheet.

You now have three tabs. Fill in the first two:

**Students**

| Roll | Name | Section |
|------|------|---------|
| 210101 | Rakib Hasan | CSE-A |
| 210102 | Sadia Islam | CSE-A |

**Courses**

| Code | Title | Section |
|------|-------|---------|
| CSE 3101 | Computer Vision | CSE-A |
| CSE 3203 | Machine Learning | CSE-B |

The **Section** column is the link between them: picking a course loads every
student whose section matches. Spelling has to match exactly.

`Attendance` fills itself — one row per student per class, with a **Class**
column so two lectures of the same course on the same day stay apart.

## 2. Set a PIN — required

The script now refuses to serve anything until a PIN exists, reads included.
In the Apps Script editor open `setPin`, change `'1234'` to your own PIN, run it
once, then change the value back to `'1234'` in the code so your real PIN isn't
sitting in the source.

Run `checkPinIsSet` afterwards and look at the execution log to confirm it
actually stored. A PIN that was never stored used to mean the sheet was wide
open; now it means nothing loads at all, which is the safer way round.

Then test it: open the `/exec` URL on its own in a private window. You should
get a refusal, not data. If you get data, you are looking at an older
deployment — see step 3.

## 3. Publish the script

**Deploy → New deployment → Web app.**

- Execute as: **Me**
- Who has access: **Anyone**

Copy the URL it gives you. It ends in `/exec`.

> "Anyone" sounds alarming but it only means the browser can call the script
> without a Google login. The script decides what it will do, and the PIN
> guards writes.

Whenever you later edit `Code.gs`, you must **Deploy → Manage deployments →
Edit → New version** for the change to go live. Editing the two data sheets
needs no redeploy.

## 4. Point the site at it

Open `index.html` and edit the three lines near the top:

```js
const API = "https://script.google.com/macros/s/AKfy..../exec";
const HEADER_TITLE = "Attendance";
const HEADER_SUB   = "Department of CSE, BAUSTK";
```

Open the file in a browser to check it loads your courses before deploying.

## 5. Deploy

**Vercel** — go to vercel.com, New Project, drag the folder in, Deploy. Done in
about a minute; you get `something.vercel.app`.

**GitHub Pages** — since you already have `nahin-jzs.github.io`, pushing this to
a repo called `attendance` and turning on Pages gives you
`nahin-jzs.github.io/attendance`. One less account to manage, and it sits
alongside your portfolio.

Either is free forever for this. Netlify and Cloudflare Pages work the same way
if you prefer them.

On your phone, open the URL and use "Add to Home Screen" — it then behaves like
an app.

---

## Using it

Five tabs across the top, all working on whichever course is selected.

### Take attendance

Everyone starts marked **Present**. Tap **A** only for the students who aren't
there, **L** for latecomers, then **Save**. The counter at the bottom updates as
you go, and the search box jumps to a roll number in a large section.

**Date + Class + Course** together identify one sitting. So you can take three
different courses on Sunday, or the same course twice on Sunday as Class 1 and
Class 2, and all three save cleanly. Save the *same* course, class and date
twice and it stops and asks whether to replace what's already there — which is
also how you correct a mistake: re-mark, save, confirm the replacement.

The Class dropdown offers four slots. Change `CLASSES_PER_DAY` near the top of
`index.html` if you need more.

### Whole term

Every student in the selected course, with present / late / absent counts and a
percentage. The header line tells you how many classes have been held and how
many students are below the threshold; those students get a red bar and a red
percentage so the list of who is at risk reads at a glance.

Two things to know:

- **Count late as present** is a checkbox, on by default. Turn it off and the
  percentages recalculate immediately, so you can see the effect before you
  decide the rule.
- The denominator is the classes *that student has a record for*, not the total
  held. Someone who transferred in during week six is judged on the classes
  they could have attended.

Set the threshold by editing `PASS_MARK` near the top of `index.html`. It ships
at 75.

Tapping any student opens their full-term record.

**Download CSV** gives you one row per student and **one column per class**, so
the sheet reads as a full grid:

| Roll | Name | 2026-01-12 C1 | 2026-01-12 C2 | 2026-01-14 C1 | … | Present | Late | Absent | Classes | Percentage |
|------|------|---------------|---------------|---------------|---|---------|------|--------|---------|------------|
| 210101 | Rakib Hasan | P | P | A | … | 28 | 1 | 3 | 32 | 88% |

`P` present, `L` late, `A` absent, blank if that student wasn't marked that day.
The totals sit at the right, so it works both as a record and as the eligibility
list the department asks for.

### One class

Pick any class you've recorded and you get that sitting's roll, grouped **absent
first**, then late, then present, with a count line at the top. The dropdown
shows newest first, and adds "Class 2" to the label only on days that actually
had more than one class of that course. Only students who were marked that day
appear, so a class taken before a student joined won't list them as missing.

**Correcting a mark:** tap P / L / A on any row. It writes to the sheet
immediately — no save button — and the count line updates as you go. The row
stays where it is rather than jumping to a different group; switch classes and
back if you want the groups redrawn.

**Deleting a class:** if you took attendance for the wrong course, wrong date or
wrong class slot, **Delete this class** removes every mark for that sitting. It
asks first and cannot be undone. The class then disappears from the term
summary, the CSV and every student's record.

### One student

A search box on its own tab. Type a roll or a name and the list narrows as you
type; tap a student and you get their whole term for that course — the running
percentage in the banner, then every class in reverse order with what they were
marked. Classes held before they joined show as "Not recorded" rather than
counting against them.

Each class row has the same P / L / A control, so this is the second way to fix
a mark: instead of hunting for the date, open the student and correct their row.
Both routes write to the same cell in the sheet.

You can also reach this from the **Whole term** list — tapping any row there
opens the same student view.

### Students

Everyone on the selected course's section, with the roll count at the top.

**Adding** takes a roll and a name and appends to the `Students` sheet — useful
when someone joins after you set the term up.

**Removing** is for students who drop or leave. It asks twice, because there are
two different things you might mean:

- **Keep the history** — they come off the roll so they stop appearing when you
  take attendance, but the classes they already attended stay in the sheet. They
  still show in the term summary, tagged *off roll*. This is usually what you
  want mid-term.
- **Erase their history** — also deletes every attendance record for that roll in
  that section, across all courses. Use this for a duplicate or a typo'd roll,
  not for a real student, since it cannot be undone.

Note that the roll is per **section**, not per course, because that's how the
`Students` sheet is organised. Removing someone takes them off every course
taught to that section.

### The raw sheet

The `Attendance` tab holds one row per student per class. Don't add formulas to
it — the script rewrites rows when you replace a day's record. If you want to
slice it a different way, do it on a separate sheet with a pivot table (Roll and
Name as rows, Status as columns, COUNTA of Status as values).

## Already running an earlier version?

1. Paste the new `Code.gs` over the old one and save.
2. Run **`upgradeSheet`** once from the function dropdown. If your `Attendance`
   sheet has no **Class** column it adds one and marks every existing row as
   class 1. If it already has one, it says so and does nothing. Running it twice
   is harmless.
3. **Deploy → Manage deployments → pencil icon → Version: New version →
   Deploy.** The URL does not change.
4. Replace `index.html` on your host, pasting your web app URL into it again.

Step 2 is not optional the first time — saving refuses to run until the Class
column exists, rather than quietly writing into the wrong columns.

The site can now delete rows as well as add them, so the PIN matters more than
it did. If you skipped it earlier, set one now.

## If something breaks

**"Could not reach the sheet"** — the deployment access isn't set to Anyone, or
the URL is missing `/exec`.

**Course list is empty** — the `Courses` tab has no rows, or the script was
deployed before you added them. Reload the site.

**"No students in this section"** — the Section text in `Students` doesn't match
the Section text in `Courses`. Watch for stray spaces.

**PIN keeps failing** — the site re-asks automatically, and shows an "Enter the
PIN" button if you dismiss the prompt. If it never asks at all, your browser is
serving a cached copy of an older `index.html`: hard-refresh with Ctrl+Shift+R
(Cmd+Shift+R on a Mac), or open the site once in a private window to confirm.

**"Could not reach the sheet" with a PIN message under it** — that was an older
build reporting a rejected PIN as a connection failure. Update `index.html`.

**The bare `/exec` URL still returns student data in a private window** — the
running deployment is older than your code. Editing `Code.gs` changes nothing
until you publish a **new version**: Deploy → Manage deployments → pencil icon →
Version: **New version** → Deploy. Then check `checkPinIsSet` actually reports a
stored PIN.

**Reports say "Nothing recorded" but the sheet has rows** — you're still on the
old deployment. See the upgrade note above.

**"The Attendance sheet is missing the Class column"** — run `upgradeSheet` once
from the Apps Script editor.

**A student shows "never marked"** — they're in the `Students` sheet but no
saved class includes them, usually because they were added after those classes
were taken.
