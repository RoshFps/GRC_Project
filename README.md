# ISMS & BCM Portfolio: Saarland University

A portfolio project in governance, risk and compliance (GRC). It builds the core documents of an information security management system (ISO/IEC 27001:2022) and business continuity management, using Saarland University as the example organisation.

I picked my own university because it is an organisation I know from the inside as a student: exam registration, the IT services, the two campuses. That makes the risks and the continuity scenario concrete instead of generic. The pieces build on each other: the risk register comes first, the Statement of Applicability links each control to the risks it treats, and the business impact analysis takes one risk from the register (R07, the campus management system failing during exam registration) and works out what an outage really costs and how fast the university has to recover.

> **Portfolio exercise.** This is not an official document of Saarland University. Only public facts are used (see the `sources.md` file in each folder); all risks, scores, owners, controls in place, implementation statuses and dates are invented.

## Contents

| Piece | Status | Folder |
|---|---|---|
| Risk register (18 risks, likelihood × impact, owner, ISO 27001 Annex A controls, remediation tracking, KPI summary and heatmap) | Done | `risk-register/` |
| Statement of Applicability (93 Annex A controls, applicability, justification, implementation status, linked risks) | Done | `soa/` |
| Business Impact Analysis, one-page continuity plan and tabletop exercise for exam registration and grade recording | Done | `bcm/` |
| Risk heatmap and KPI dashboard (Python, reads the register; static HTML page plus PNG charts) | Done | `dashboard/` |

## Risk register

- `risk-register/saarland-university-risk-register.xlsx`: the register. Sheets: Read Me, Risk Register, Summary, Scoring Method, Lists, Sources.
- `risk-register/build_register.py`: generates the workbook and the sources list (`pip install openpyxl`, then `python3 build_register.py out.xlsx sources.md`).
- `risk-register/sources.md`: public sources, standards and incident reports behind the facts used.

## Statement of Applicability

- `soa/saarland-university-soa.xlsx`: all 93 ISO/IEC 27001:2022 Annex A controls with paraphrased titles, applicability, reasons for inclusion (risk, legal, contractual, baseline), justification, implementation status, owner, evidence and the register risks each control treats. Sheets: Read Me, SoA, Summary, Lists, Sources.
- `soa/build_soa.py`: generates the workbook and its sources list, reading the related risks from the risk register (`python3 build_soa.py ../risk-register/saarland-university-risk-register.xlsx out.xlsx sources.md`).
- `soa/sources.md`: standards and references behind the SoA.

## Business continuity

- `bcm/saarland-university-bia.xlsx`: business impact analysis for exam registration and grade recording in LSF. Impact over time in five categories for peak and off-peak weeks, MTPD, RTO and RPO with reasoning, dependencies, linked register risks and actions to close the gaps. Result: in the registration window the process can be down for at most 2 days, so the target is back within 24 hours with at most 1 hour of lost data. Today's nightly backup and roughly 3-day rebuild miss both.
- `bcm/build_bia.py`: generates the workbook (`python3 build_bia.py out.xlsx`).
- `bcm/continuity-plan.md`: one-page plan: when to activate, who does what, the paper fallback, deadline extensions and the return to normal.
- `bcm/tabletop-exercise.md`: "Deadline Day", a two-hour exercise where LSF is hit by ransomware on the last two days of the winter registration window, with six injects, expected responses and how to score them.
- `bcm/sources.md`: public facts (LSF, iTAN, semester dates) and standards (ISO 22301, BSI-Standard 200-4) behind the documents.

## Risk dashboard

`dashboard/risk_dashboard.py` reads the risk register and builds a one-page report on the current risk position:

- KPIs: risks on the register, open risks, overdue remediations, high or critical risks before and after treatment, total score reduction.
- A 5×5 heatmap of inherent and residual risk side by side.
- Inherent against residual score for every risk, with the risk appetite line.
- Risks per owner, split by remediation status.
- A table of overdue remediations and the full register.

Outputs in `dashboard/output/`: `risk-dashboard.html` (a single page with the charts inline; open it in a browser) and `heatmap.png`, `movement.png`, `owners.png` for slides.

```
pip install openpyxl matplotlib
cd dashboard
python3 risk_dashboard.py ../risk-register/saarland-university-risk-register.xlsx
python3 risk_dashboard.py ../risk-register/saarland-university-risk-register.xlsx --as-of 2026-10-08 --out output
```

`--as-of` sets the reporting date used for overdue checks (default: today). Scores, ratings and overdue flags are recalculated from likelihood, impact, status and due date, and the rating thresholds come from the Scoring Method sheet, so editing the register and re-running is enough. On first run the script downloads Source Serif 4 and IBM Plex from Google Fonts into `dashboard/.fonts/`; offline it falls back to the default font.

## How to use it

Open the Excel files directly; every workbook starts with a Read Me sheet. Yellow cells are inputs and can be changed, the rest recalculates. To regenerate any workbook, install `openpyxl` and run the build script in its folder.
