# ISMS & BCM Starter Kit: Saarland University

A portfolio project in governance, risk and compliance (GRC). It builds the core documents of an information security management system (ISO/IEC 27001:2022) and business continuity management, using Saarland University as the example organisation.

> **Portfolio exercise.** This is not an official document of Saarland University. Only public facts are used (see `risk-register/sources.md` and `soa/sources.md`); all risks, scores, owners, controls in place, implementation statuses and dates are invented.

## Contents

| Piece | Status | Folder |
|---|---|---|
| Risk register (18 risks, likelihood × impact, owner, ISO 27001 Annex A controls, remediation tracking, KPI summary and heatmap) | Done | `risk-register/` |
| Statement of Applicability (93 Annex A controls, applicability, justification, implementation status, linked risks) | Done | `soa/` |
| Business Impact Analysis, continuity plan and tabletop exercise | Planned | |
| Python heatmap and KPI dashboard reading the register | Planned | |

## Risk register

- `risk-register/saarland-university-risk-register.xlsx`: the register. Sheets: Read Me, Risk Register, Summary, Scoring Method, Lists, Sources.
- `risk-register/build_register.py`: generates the workbook and the sources list (`pip install openpyxl`, then `python3 build_register.py out.xlsx sources.md`).
- `risk-register/sources.md`: public sources, standards and incident reports behind the facts used.

## Statement of Applicability

- `soa/saarland-university-soa.xlsx`: all 93 ISO/IEC 27001:2022 Annex A controls with paraphrased titles, applicability, reasons for inclusion (risk, legal, contractual, baseline), justification, implementation status, owner, evidence and the register risks each control treats. Sheets: Read Me, SoA, Summary, Lists, Sources.
- `soa/build_soa.py`: generates the workbook and its sources list, reading the related risks from the risk register (`python3 build_soa.py ../risk-register/saarland-university-risk-register.xlsx out.xlsx sources.md`).
- `soa/sources.md`: standards and references behind the SoA.
