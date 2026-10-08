# Tabletop Exercise: "Deadline Day"

**Saarland University · Tests the continuity plan for exam registration and grade recording · Planned for December 2026**

> Portfolio exercise. Not an official university document. The scenario, times and figures are invented.

## Purpose

Check whether the people named in `continuity-plan.md` can keep exam registration going when LSF fails on the last two days of the winter registration window, and whether the BIA targets (RTO 24 hours, RPO 1 hour at peak) are realistic with today's backups.

## Objectives

1. Activate the plan within the agreed time and run the first crisis call.
2. Decide on manual fallback and deadline extensions within 4 hours of the first report.
3. Handle a suspected cyber attack without restoring an infected system.
4. Agree on what to tell students, and when.
5. Find out how much registration data would really be lost.

## Format

- **Duration:** 2 hours, discussion only, no systems touched
- **Participants:** Head of Student Services (plan lead), Head of HIZ, CISO, Vice President for Teaching (or deputy), two exam office leads, Press and Communications, Data Protection Officer
- **Facilitator:** Information security team; one note taker
- **Rules:** Answer as you would on the day. "I would check the plan" is fine if you say what it tells you. No blame.

## Scenario and injects

The winter registration window closes on **Friday 22 January 2027 at 23:59** (assumed date). About 40% of all registrations usually arrive in the last three days.

| Time (exercise) | Inject | Questions for the group |
|---|---|---|
| **Thu 21 Jan, 07:45** | The HIZ service desk gets calls: the HIS-Portal shows an error page. Two exam offices report the same. Monitoring shows the LSF database server is not responding. | Who is told first? Does this meet the activation criteria? Who opens the crisis call, and who is on it? |
| **08:30** | HIZ finds the database files renamed with an unknown extension and a text file asking for payment. Two other file servers show the same. | Who now leads: the plan lead or the CISO? What is isolated? Can HIZ restore from last night's backup, and is that backup safe? |
| **09:30** | The nightly backup ran at 02:00 on the same storage system. Its status is unknown. Restore estimate if the backup is clean: 3 days. | Compare with RTO 24 hours and MTPD 2 days. Do you switch to the manual fallback now? What data has been lost since 02:00? |
| **11:00** | Students post screenshots on social media; the student council asks for a statement. A regional newspaper calls Press and Communications. | What do you say to students and to the press? Who approves the text? Is a deadline extension announced now or later? |
| **14:00** | Exam offices report long queues and 400 paper forms already received. One faculty has no paper form ready. Some students registered on Wednesday evening say their registrations are gone. | How are paper registrations recorded and later entered? How do you prove which online registrations existed before the failure? |
| **Fri 22 Jan, 10:00** | HIZ reports that Thursday's 02:00 backup was encrypted too. Wednesday's backup is clean and the restore is under way, finished Sunday evening. Logs suggest the attacker had access for 10 days. | Is personal data exposed? Does the Data Protection Officer report to the supervisory authority within 72 hours? When is the new deadline? How do you check registrations and grades before reopening? |

## Expected good responses

- Plan activated before 09:45 (2 hours after the first report).
- No restore before the CISO confirms the backup is clean and the attack path is closed.
- Manual fallback and a uniform deadline extension (outage plus 2 days) decided by 11:45.
- One message to students from one source, repeated on website, email and social media.
- Data Protection Officer involved by 10:00 Friday at the latest; decision on GDPR Art. 33 report documented.
- The group notices that with only nightly backups on the same storage, up to 30 hours of registrations can be lost (Wednesday 02:00 to Thursday 07:45), far above the 1-hour RPO.

## Evaluation

The note taker records, for each inject: decision taken, time, who decided, and any gap in the plan. Score each objective:

| Rating | Meaning |
|---|---|
| Met | Done in time, following the plan |
| Partly met | Done, but late or outside the plan |
| Not met | Not done, or the plan did not cover it |

## After the exercise

Within 2 weeks the facilitator writes a short report with findings and actions, each with an owner and due date. Findings update `continuity-plan.md`, the BIA and the risk register (mainly R07, R01, R12 and R15).
