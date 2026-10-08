# Sources for the Saarland University Statement of Applicability

Portfolio exercise, not an official university document. Sources checked on 8 October 2026. Control titles are paraphrased; all statuses, owners, documents and the scope are invented.

## Standard

- [ISO/IEC 27001:2022 Information security management systems: Requirements (with Amd 1:2024)](https://www.iso.org/standard/27001)  
  Supports: Clause 6.1.3 d) requires a Statement of Applicability; Annex A lists the 93 controls in four themes.  
  Used in: Whole SoA; control numbers and themes. Access: Paid standard (about CHF 155), no copy included.
- [ISO/IEC 27002:2022 Information security controls](https://www.iso.org/standard/75652.html)  
  Supports: Purpose and guidance for each control, used to write the paraphrased titles and justifications.  
  Used in: Columns B and I. Access: Paid standard, no copy included.

## Law

- [EU General Data Protection Regulation (GDPR), Regulation (EU) 2016/679](https://eur-lex.europa.eu/eli/reg/2016/679/oj)  
  Supports: Art. 32 security of processing; Art. 33 breach notification within 72 hours.  
  Used in: Reason L; A.5.26, A.5.34, A.8.10, A.8.11 and others. Access: Free.
- [Basic Law for the Federal Republic of Germany (Grundgesetz), Art. 5(3)](https://www.gesetze-im-internet.de/gg/art_5.html)  
  Supports: Freedom of science, research and teaching, the reason web filtering is limited to known malicious sites.  
  Used in: A.8.23. Access: Free.

## Internal

- [Saarland University risk register (this project)](../risk-register/saarland-university-risk-register.xlsx)  
  Supports: Risks R01 to R18 and the Annex A controls that treat them.  
  Used in: Column M (related risks), reason R. Access: Included in this repository.

## Files in this folder

- `saarland-university-soa.xlsx`: the SoA (sheets Read Me, SoA, Summary, Lists, Sources).
- `build_soa.py`: Python script that generates the workbook and this file (`python3 build_soa.py ../risk-register/saarland-university-risk-register.xlsx out.xlsx sources.md`).
- `sources.md`: this list.

The ISO standards are paid and cannot be shared. The GDPR and the Basic Law are free to read at the links above.
