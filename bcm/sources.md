# Sources for the BIA, continuity plan and tabletop exercise

Portfolio exercise, not an official university document. Sources checked on 8 October 2026. Only the facts below come from public sources; impact scores, MTPD, RTO and RPO, volumes, recovery times, roles and the exercise scenario are invented.

## Organisation and process

- [Saarland University (ZPL): Fragen und Antworten A bis Z](https://www.uni-saarland.de/en/einrichtung/zpl/studienorganisation/abisz.html)  
  Supports: Exam registration and checking of results run in LSF; the deadline is shown at each exam; late registration is not possible where registration is mandatory.  
  Used in: BIA (Process, Impact Over Time); continuity plan section 1.
- [Saarland University, Systems Engineering: Prüfungsanmeldung und -termine](https://www.uni-saarland.de/fachrichtung/systems-engineering/studium/pruefungsanmeldung-und-termine.html)  
  Supports: Online registration with the university account, confirmed by iTAN; results and registration status in the HIS-Portal; if online registration fails, students contact the exam office and register in writing within the deadline.  
  Used in: BIA (Process, RTO & RPO); continuity plan step 4.
- [Saarland University: Semestertermine Studienjahr 2026/2027 (PDF)](https://www.uni-saarland.de/fileadmin/upload/dezernat/ls/semestertermine/Studienjahr2026-2027.pdf)  
  Supports: Winter semester 2026/27 lecture period 12 October 2026 to 5 February 2027, semester end 31 March 2027.  
  Used in: BIA (Process, peak periods). The registration window dates are assumed.
- [Wikipedia: Saarland University](https://en.wikipedia.org/wiki/Saarland_University)  
  Supports: About 16,300 students, used to estimate registration volume.  
  Used in: BIA (Process, volume assumption).

## Standards and method

- [ISO 22301:2019 Business continuity management systems: Requirements](https://www.iso.org/standard/75106.html)  
  Supports: Business impact analysis, MTPD, RTO and RPO (clause 8.2.2); exercising (clause 8.5). Paid standard, no copy included.
- [BSI-Standard 200-4 Business Continuity Management](https://www.bsi.bund.de/DE/Themen/Unternehmen-und-Organisationen/Standards-und-Zertifizierung/IT-Grundschutz/BSI-Standards/BSI-Standard-200-4-Business-Continuity-Management/bsi-standard-200-4_Business_Continuity_Management_node.html)  
  Supports: German BCM method, used as background for the BIA approach. Free download.
- [ISO/IEC 27001:2022](https://www.iso.org/standard/27001)  
  Supports: Annex A controls A.5.29 Information security during disruption and A.5.30 ICT readiness for business continuity. Paid standard.

## Law

- [EU General Data Protection Regulation, Art. 33](https://eur-lex.europa.eu/eli/reg/2016/679/oj)  
  Supports: 72-hour breach notification, used in the plan and the exercise.

## Files in this folder

- `saarland-university-bia.xlsx`: the BIA (sheets Read Me, Process, Impact Over Time, RTO & RPO, Dependencies, Linked Risks, Actions, Impact Scale, Sources).
- `build_bia.py`: generates the workbook (`pip install openpyxl`, then `python3 build_bia.py out.xlsx`).
- `continuity-plan.md`: one-page continuity plan.
- `tabletop-exercise.md`: exercise scenario with injects and evaluation.
- `sources.md`: this list.
