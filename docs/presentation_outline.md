# Personal Toolbox — First Introduction to Regional HR Teams

**Speaker script & slide outline.** Audience: HR colleagues from other regions who know
recruitment and people operations well, but do not think in terms of systems or software.
Goal: they leave convinced the app is (1) genuinely efficient, (2) carefully built, and
(3) heading somewhere — one central, systematic HR data platform.

> **Ghi chú cho người trình bày (không đưa lên slide)**
> - Tổng 35–40 phút: 8 phút phần 1–3, 15 phút demo tuyển dụng, 7 phút nhân sự/đào tạo,
>   5 phút "độ chỉnh chu" + roadmap, còn lại Q&A.
> - Nguyên tắc kể chuyện: **mỗi slide bắt đầu bằng việc HR phải làm, không bắt đầu bằng
>   tính năng**. Người nghe không quan tâm "SQLite", họ quan tâm "không phải gõ lại".
> - Chỗ nào có `[…]` là số liệu của bạn — điền trước khi trình bày, đừng ước lượng trên sân.
> - Tránh hẳn các từ: database schema, migration, repository, SQLite, thread. Dùng:
>   "one shared list", "the app remembers", "runs in the background".

---

## Slide 1 — Title

**Personal Toolbox**
One app for the HR work we repeat every week.

- Presenter · Region · Date
- Sub-line: *Recruitment · People data · Training · Daily office tasks*

**Say:** "What you'll see is not a prototype. It is the tool I use for my own daily HR
work, and today is the first time we show it outside our team. I'd like your reaction on
two things: what would help you most, and what we should standardise across regions."

---

## Slide 2 — Where HR data lives today

Four things that are true in almost every HR team:

| What happens | What it costs us |
|---|---|
| CVs sit in folders, named however the sender named them | Nobody can find last year's candidate for a similar role |
| Candidate tracking lives in an Excel file, copied per batch, per person | Two people, two versions, no single truth |
| Interview feedback arrives by email, chat, or on paper | The evaluation is lost the moment the hire is closed |
| The same person is typed again into the employee file after hiring | Same data, entered three times, three chances to be wrong |

**Say:** "None of this is anyone's fault — it is what happens when the work is faster than
the filing. The app does not ask you to work differently. It removes the retyping."

---

## Slide 3 — What it is, in one picture

- A desktop app on your own Windows laptop. One window, a menu on the left.
- **10 tools in 5 groups**: Recruitment · Human Resources · Files & Documents · Office ·
  Master Data.
- No server to buy, no browser login, no IT installation ticket — a single program file.
- Everything you enter is kept in **one place** on your machine, not in scattered files.
- Works offline. Only the AI features need internet.

**Say:** "Think of it as one drawer instead of fifteen. Same drawer for candidates,
employees, courses, and the small weekly chores."

*(Screenshot: home screen with the tool cards.)*

---

## Slide 4 — Section divider: RECRUITMENT

*"From a folder of PDFs to a shortlist — without typing a candidate twice."*

---

## Slide 5 — One recruitment run, end to end

The demo you are about to see follows one real sequence:

1. **Receive** — a folder of CV files
2. **Name** — file names normalised into batch codes
3. **Read & score** — AI reads each CV against the position's job description
4. **Shortlist** — search and filter one candidate list
5. **Invite** — interview invitations sent through Outlook, with the CV attached
6. **Evaluate** — feedback captured per interviewer, per round
7. **Report** — the Excel file you already know, produced in one click

**Say:** "Seven steps. Today, six of them involve retyping. In the app, data is entered
once at step 3 and reused all the way to step 7."

---

## Slide 6 — Step 1–2 · CV intake

- Point the app at the folder of CVs.
- **Normalize file names**: choose a prefix and a starting number, list the words you want
  removed from candidate names, then preview a before/after table — and correct any name by
  hand — before renaming the whole batch.
- Result: `[PREFIX][number]_Candidate Name.pdf`, one consistent naming rule for everyone.

**Why it matters:** the batch code stays with the candidate from the file name to the final
Excel report, so a CV can always be traced back to the file it came from.

---

## Slide 7 — Step 3 · AI reads the CV and scores the fit

Select the **position** — the app already knows which job description belongs to it, so
nobody hunts for a JD file. Add extra instructions for the AI if you want ("must have
electronics manufacturing background").

For every CV in the folder, the app captures:

- Full name · date of birth · email · phone
- **Fit score, 0–100**, plus a short written justification
- **Strengths** and **weaknesses** against that specific job description
- The CV file itself, linked to the record

Duplicates are caught on the spot: same email or phone as an existing candidate, and the
app asks you whether to overwrite or export separately — it never quietly merges two people.

**Say:** "This is the step that pays for the whole app. Reading and summarising 40 CVs is an
evening of work. Here it runs in the background with a progress bar while you do something
else, and the result is not a note in your head — it is data you can search next year."

*(Fill in: `[your own number]` CVs scanned in `[minutes]` during the last opening.)*

---

## Slide 8 — Step 4 · One candidate list, searchable

- Free-text search across the whole candidate record.
- Filters that match how we actually talk: **Position · Department · Status · Batch**.
- Click the CV file name in the row to open the PDF. If the file was moved or renamed, the
  app asks you to point at it once and then remembers the new location.
- Tick several candidates and act on all of them at once.
- Right-click a row for **Quick edit** (interview feedback) and **Update source** — the job
  board or referral the CV came from, filled in for the whole selection in one go.

**Why it matters:** the question "who did we interview for this role last year, and why
didn't we hire them?" becomes a search instead of an archaeology project.

---

## Slide 9 — Step 5 · Interview invitations through Outlook

Tick the candidates you want to invite, choose the round (1st, 2nd, 3rd):

- The app picks the **email template attached to that position for that round** — so the
  wording is consistent no matter who sends it.
- You set a time per candidate; each following candidate is pre-set to start right after
  the previous one, and *Skip* passes over anyone you'd rather schedule later.
- It then opens one **Outlook meeting window per candidate**, pre-filled: recipient, CC,
  time, body text from the template, **and the candidate's CV attached**.
- You review each window, add the meeting room, and press Send. Outlook does what Outlook
  does best — invitation, calendar entry, room booking.
- Straight after, the app offers to move exactly those candidates to the next status.

For **Application Thank You** letters, it opens a normal mail instead — no time slots, no
status change.

**Say:** "Notice what the app deliberately does *not* do: it never sends anything by itself.
Every mail stops in front of you first. That was a design decision, not a limitation."

---

## Slide 10 — Step 6 · Interview feedback, per interviewer

The full candidate form has around thirty fields. The thing you do most often after an
interview needs five. So there is a short form for exactly that:

- Status · result · phone screen date and note
- For each of the three rounds: date, overall result, and a list of interviewers —
  **add as many interviewers as attended**, each with their own conclusion and comments.
- Interviewer names come from the **employee directory**, so title and department are looked
  up automatically and never retyped.
- Dates open a calendar, show `dd/MM/yyyy`, and may be left empty — a round that hasn't
  happened yet does not get a meaningless date.

**Why it matters:** feedback is stored per person, not merged into one paragraph. In the
Excel report, each comment still carries the name of who wrote it.

---

## Slide 11 — Step 7 · The Excel report you already use

One click on **Export to Excel** produces the familiar **Candidates** sheet — 25 columns,
`A` to `Y`: batch, ID, name, position, source, contact, score, AI evaluation, status,
result, phone screen, and four columns for each of the three interview rounds.

- Dropdown lists on the columns that should be chosen, not typed.
- Status and result colour-coded; duplicate email addresses highlighted in red.
- If the file already exists, the app asks: **Append · Overwrite · Cancel**. It never
  silently changes a file you already sent to someone.
- `[166]` records export in **under a second**, and the file is about **56 KB** — no
  formulas, no links to other files, so it opens instantly on any machine.

**Say:** "Your reporting does not have to change at all. The Excel file stays. It is simply
generated instead of maintained."

---

## Slide 12 — Section divider: PEOPLE & TRAINING

*"The same principle applied after the hire: type it once."*

---

## Slide 13 — The employee master file

- Around **90 fields**, matching our *Master HC file* almost column for column.
- **Bulk import from Excel matches by column heading, not column position** — reorder the
  columns or add a new one, and the import still lands correctly.
- Emergency contact is one messy cell in the source file (name and phone number mixed, in
  every possible format); the app splits it into a name and a phone number on import.
- Working / resigned is derived from the termination date, so it can never disagree with it.
  Resigned colleagues are hidden by default; one tick shows them.
- 90 columns would be unreadable, so the table shows a sensible default set and you switch
  the rest on in **Columns**, grouped as Identity · Contact · Education · and so on. Your
  choice is remembered.
- Pre-loaded directories so everyone chooses from the same list, not free text:
  **20 departments · 6 employee types · 12 job levels · 42 cost centers**.

---

## Slide 14 — New joiners: the application form reads itself

The **Import application form** button takes the DLVN Application Form as the new employee
submitted it — and handles both realities:

| How it arrives | How the app handles it |
|---|---|
| Typed into the template file | Reads the cells the employee filled in |
| Printed, filled in by hand, scanned to PDF | Reads the handwriting from the scanned pages |

- Several forms can be processed in one go.
- **Every form opens a review screen, pre-filled, before anything is saved.** Handwriting is
  not always legible, so a human confirms. *Save* writes it; *Cancel* skips that form.
- Before saving it checks employee code, ID card number, and full name against existing
  records, and shows you any possible duplicate.
- Fields the form has but the employee record does not (references, expected salary…) are
  deliberately not stored — we keep what we actually use.

---

## Slide 15 — Training records

- **Enroll** employees onto a course from the employee list.
- **Course Manager** shows enrolments and learning status.
- The attendance sheet problem: we print the roster, everyone signs it by hand, and the
  whole stack is scanned into one PDF. The app splits that PDF page by page, reads which
  rows carry a handwritten signature, matches them by **employee code**, and marks those
  people **Completed**.

**Say:** "That is the clearest example of what this app is for: nobody should be paid to
tick a hundred boxes that a scan already answers."

---

## Slide 16 — The small things, every week

| Tool | What it does |
|---|---|
| **Gate-Open Mail** | Every morning it checks today's Outlook calendar for interviews and drafts the mail asking Security to open the gate — shown to you first |
| **Interview Follow-up** | Scans the last month of interviews and reminds you which candidates are still waiting for an answer |
| **Birthday emails** | Sends the birthday card to the right employees on the right day |
| **Split Payroll** | Splits one master payroll file into a separate file per vendor |
| **Quarterly Bonus** | Aggregates a quarter's attendance issues into one summary sheet for bonus review |
| **PDF → Text** | Turns a PDF back into text that keeps its layout, ready to paste into Word |

**Say:** "Individually, five minutes each. Together, that is the reason my inbox closes on
time. And every one of them started as something I was doing by hand."

---

## Slide 17 — Built to be trusted

Six rules the app follows everywhere. This is the part I'd ask you to judge it on:

1. **AI never writes without a human.** Every AI result — CV score, application form,
   signature sheet — lands in a screen you approve first.
2. **Nothing is overwritten silently.** Duplicate candidates, duplicate employees, existing
   Excel files: the app asks, and *Cancel* is always one of the answers.
3. **History is added, never replaced.** Each AI evaluation is kept with its date, the model
   used, and which version of the CV it read. You can see how an assessment changed.
4. **Machine-written columns cannot be edited by hand.** Batch, ID, score, AI evaluation —
   editing those would corrupt the evaluation history, so the app doesn't allow it.
5. **Shared lists instead of free typing.** Departments, levels, positions, statuses, email
   templates live in one directory used by every screen.
6. **Small details, on purpose.** Dates always `dd/MM/yyyy` and allowed to be empty; the
   scroll wheel can't accidentally change a date; phone numbers keep their leading zero in
   Excel; long jobs run in the background with a progress bar instead of freezing.

**Say:** "Efficiency is easy to demo. Trust is what decides whether a tool survives its
first month, so most of the work went here."

---

## Slide 18 — Where AI is used, and where it isn't

| Used for | Not used for |
|---|---|
| Reading a CV and scoring fit against the JD | Deciding who to hire or reject |
| Reading a filled-in application form | Writing anything into the database unreviewed |
| Reading handwritten signatures on a scanned roster | Sending any email on its own |
| Reading scanned PDFs back into text | Anything while you are offline |

- What leaves the machine: only the specific file being read, sent to Google Gemini for
  that request. Nothing is uploaded in the background, and nothing is stored outside the app.
- The AI key is set once in **Settings** and shared by all AI features.
- If AI is unavailable, every screen still works — you type instead.

**Say:** "I'd rather answer this before you ask it. The AI is a fast reader here. It has no
authority over any decision."

---

## Slide 19 — Today → where this is going

**Today.** One app per person, data kept on that laptop. Fast, private, zero setup — but
each of us still has our own copy.

**Next: one shared source of truth.**

| Phase | What changes | What you get |
|---|---|---|
| **Now** | Individual use, one HR team | Time back on CV screening, invitations, reporting |
| **Next** | Shared central database, several regions on the same lists | One candidate history, one employee master, no reconciliation |
| **Then** | Reporting on top of that data — headcount, hiring funnel, time-to-hire, training coverage — plus access rights per role | Regional numbers that agree with each other, produced automatically |

**Say:** "Everything you saw was built so this step is possible: consistent lists, records
that keep their history, one place per fact. Moving from one laptop to one shared database
is then a change of location — not a rebuild."

---

## Slide 20 — What I'd like from you

1. **One pilot volunteer per region** — use it on a real hiring batch and tell me where it
   annoys you.
2. **Confirm the report columns.** If your region needs a column the Excel sheet doesn't
   have, now is the cheapest moment to add it.
3. **Agree the shared lists.** Candidate statuses, job levels, sources, departments — a
   central database is only useful if we all name things the same way.
4. **Tell me your biggest repeated manual task.** Every tool in this app started as one
   sentence like that.

---

## Slide 21 — Close

**Personal Toolbox** — the work is the same. The typing isn't.

Contact · how to get a copy · where the pilot starts.

---

## Appendix — Q&A preparation

| Likely question | Short answer |
|---|---|
| "Is candidate data safe?" | It stays on your machine. Only the file being read is sent for AI reading, per request. No background upload. |
| "Can it be wrong?" | Yes — AI reading can misread, especially handwriting. That is why every AI result is reviewed before it is saved, and machine-generated scores are kept with their date and source. |
| "Do we have to stop using Excel?" | No. The Excel report is generated by the app, same columns, in under a second. |
| "What if two of us edit the same candidate?" | Today each person has their own copy — that is exactly the limitation the central database phase removes. |
| "Who maintains it?" | Built and maintained in-house. New tools are added one file at a time; a request usually turns into a working screen quickly. |
| "Does it need IT to install?" | No admin rights, no separate components to install. One program file, plus Outlook and Excel which we already have. |
| "How much does the AI cost?" | It runs on a Gemini key set in Settings; the free tier covers ordinary daily volume, with paid capacity available if we scale up regions. |
| "Can it read CVs in our language?" | The AI reads mixed-language CVs. The app interface is English everywhere. |

---

## Appendix — Demo checklist (chuẩn bị trước)

- [ ] A prepared folder of `[5–8]` sample CVs (anonymised), plus one position with its JD attached.
- [ ] Outlook open and logged in; a safe test recipient (yourself) for the invitation demo.
- [ ] One scanned application form and one signed attendance sheet ready to open.
- [ ] An existing Excel export on the desktop, to demo the Append / Overwrite / Cancel prompt.
- [ ] Internet checked — run one AI scan before the room fills up, so the quota and key are proven.
- [ ] A fallback: screen recording of the AI scan, in case the network or quota fails live.
- [ ] Window maximised, app font size readable from the back row.
