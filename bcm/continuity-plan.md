# Continuity Plan: Exam Registration and Grade Recording

**Saarland University · Process owner: Head of Student Services · Version 0.1, October 2026**

> Portfolio exercise. Not an official university document. Public facts are listed in `sources.md`; roles, figures and procedures are invented.

## 1. Scope and targets

Covers exam registration, withdrawal, attendance lists and grade entry in the LSF exam administration (HIS-Portal) for both campuses. Targets from the BIA (`saarland-university-bia.xlsx`):

| | Peak (registration window, grading deadline) | Off-peak |
|---|---|---|
| Maximum tolerable disruption (MTPD) | 2 days | 4 weeks |
| Recovery time objective (RTO) | 24 hours | 1 week |
| Recovery point objective (RPO) | 1 hour | 24 hours |

## 2. When to activate

Activate if, **during a peak period**, students cannot register or examiners cannot enter grades for **more than 2 hours**, or if any data loss or tampering in the exam database is suspected. Off-peak, activate if the expected outage is longer than 2 working days. The HIZ service desk or any exam office reports to the **Head of Student Services**, who activates the plan. If the cause is a cyber attack (risk R01), the CISO's incident response runs in parallel and decides on isolation first.

## 3. Roles

| Role | Task |
|---|---|
| Head of Student Services (plan lead) | Activates the plan, chairs the crisis call, decides on fallback |
| Head of HIZ | Restores LSF and the database; gives recovery estimates every 2 hours |
| CISO | Leads incident response if an attack is suspected; decides when systems are clean |
| Vice President for Teaching | Asks exam board chairs to extend deadlines; signs the announcement |
| Exam office leads (each faculty) | Run the paper fallback; keep a list of every manual registration |
| Press and Communications | Website banner, student email, social media |
| Data Protection Officer | Assesses whether lost or exposed data must be reported (GDPR Art. 33, 72 hours) |

## 4. Response steps

**First hour**
1. Plan lead opens a crisis call with HIZ, CISO and exam office leads.
2. HIZ confirms the cause and gives a first recovery estimate. Do not restore if an attack is suspected until the CISO agrees.
3. Website banner and email to students: "Exam registration is down. Your deadline is protected. Do not register twice."

**Within 4 hours**

4. If the estimate exceeds the RTO, start the **manual fallback**: exam offices accept written registration (paper form or signed scan by email) with a timestamp, as the exam pages already allow when online registration fails.
5. Vice President for Teaching asks all exam board chairs to extend registration deadlines by the length of the outage plus 2 days, using the pre-agreed rule.
6. Exam offices export the last attendance lists available, or rebuild them from the last backup.

**Until recovery**

7. HIZ restores LSF from the latest clean backup or switches to the standby site; updates every 2 hours.
8. Plan lead checks the backlog daily against the MTPD.

## 5. Return to normal

- Before reopening, exam offices check a sample of registrations and grades against their own records, to detect data lost after the last backup (RPO) or changed by an attacker.
- Manual registrations are entered with the original timestamp and confirmed to each student by email.
- Announce the end of the outage and the final new deadlines.
- Within 2 weeks: lessons learned meeting, update this plan, the BIA and risk R07 in the register.

## 6. Keep the plan ready

Paper forms and the import template are stored with each exam office (BIA action A5). The plan is tested once a year with the tabletop exercise in `tabletop-exercise.md`, before the winter registration window. Contact details are kept in a separate, offline contact list (not part of this portfolio).
