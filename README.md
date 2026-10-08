# ISMS & BCM Starter Kit: Saarland University

A portfolio project in governance, risk and compliance (GRC). It builds the core documents of an information security management system (ISO/IEC 27001:2022) and business continuity management, using Saarland University as the example organisation.

> **Portfolio exercise.** This is not an official document of Saarland University. Only public facts are used (see `risk-register/sources.md`); all risks, scores, owners, controls in place and dates are invented.

## Contents

| Piece | Status | Folder |
|---|---|---|
| Risk register (18 risks, likelihood × impact, owner, ISO 27001 Annex A controls, remediation tracking, KPI summary and heatmap) | Done | `risk-register/` |
| Statement of Applicability (93 Annex A controls) | Planned | |
| Business Impact Analysis, continuity plan and tabletop exercise | Planned | |
| Python heatmap and KPI dashboard reading the register | Planned | |

## Risk register

- `risk-register/saarland-university-risk-register.xlsx`: the register. Sheets: Read Me, Risk Register, Summary, Scoring Method, Lists, Sources.
- `risk-register/build_register.py`: generates the workbook and the sources list (`pip install openpyxl`, then `python3 build_register.py out.xlsx sources.md`).
- `risk-register/sources.md`: public sources, standards and incident reports behind the facts used.
