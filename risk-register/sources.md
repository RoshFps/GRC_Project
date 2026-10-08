# Sources for the Saarland University risk register

Portfolio exercise, not an official university document. Sources checked on 8 October 2026. Only the facts below come from public sources; all risks, scores, owners, existing controls, statuses and dates in the register are invented.

## Organisation

- [Saarland University: My new working environment (Saarbrücken and Homburg campuses)](https://uni-saarland.de/en/page/newly-appointed-professors/your-professorship/working-environment.html)  
  Supports: Campuses in Saarbrücken and Homburg; Faculty of Medicine at the Medical Center in Homburg; University IT Centre (HIZ) with HR and Finance at the Meerwiesertalweg site; IT Service Desk in the Campus Center.  
  Used in: Read Me; R03, R07, R09. Access: Free web page.
- [Saarland University: Welche weiteren Anlaufstellen gibt es an der UdS?](https://www.uni-saarland.de/page/neuberufen/persoenliches/anlaufstellen.html)  
  Supports: HIZ is the shared IT service provider of htw saar and Saarland University.  
  Used in: Read Me; owner role 'Head of University IT Centre (HIZ)'. Access: Free web page.
- [Hochschul-IT-Zentrum Saar (HIZ) website](https://www.hiz-saarland.de/start/)  
  Supports: Official site of the IT centre (listed for reference; its content was not used).  
  Used in: Owner role. Access: Free web page.
- [Saarland University: The Homburg campus](https://uni-saarland.de/en/studies/saarland/the-homburg-campus.html)  
  Supports: Medical campus in Homburg.  
  Used in: R03. Access: Free web page.
- [Student council computer science: Internet @ Uni](https://cs.fs.uni-saarland.de/en/special/internet/)  
  Supports: University offers eduroam WiFi and a VPN for remote access.  
  Used in: R02, R13. Access: Free web page.
- [Wikipedia: Saarland University](https://en.wikipedia.org/wiki/Saarland_University)  
  Supports: Background only: founded 1948, six faculties, about 16,300 students.  
  Used in: Background, not in the register. Access: Free web page.
- [DFN: The national research and education network](https://dfn.de/en/network)  
  Supports: German research network (X-WiN). A Saarland University uplink via DFN is assumed, not confirmed by a university page.  
  Used in: R05 (assumption). Access: Free web page.

## Standard

- [ISO/IEC 27001:2022 Information security management systems: Requirements (with Amd 1:2024)](https://www.iso.org/standard/27001)  
  Supports: Annex A with 93 controls; control numbers and names in column L.  
  Used in: Column L; Statement of Applicability later. Access: Paid standard (about CHF 155), no copy included.
- [ISO/IEC 27002:2022 Information security controls](https://www.iso.org/standard/75652.html)  
  Supports: Guidance for each Annex A control.  
  Used in: Column L. Access: Paid standard, no copy included.

## Method

- [BSI-Standards overview (incl. BSI-Standard 200-3 risk management)](https://www.bsi.bund.de/DE/Themen/Unternehmen-und-Organisationen/Standards-und-Zertifizierung/IT-Grundschutz/BSI-Standards/bsi-standards_node.html)  
  Supports: German government method for risk analysis, used as background. The 5x5 scales here are this exercise's own, not BSI's.  
  Used in: Scoring Method. Access: Free download from BSI.
- [BSI-Standard 200-3 English PDF](https://www.bsi.bund.de/SharedDocs/Downloads/EN/BSI/Grundschutz/International/bsi-standard-2003_en_pdf.html)  
  Supports: Same, English version.  
  Used in: Scoring Method. Access: Free download from BSI.

## Law

- [EU General Data Protection Regulation (GDPR), Regulation (EU) 2016/679](https://eur-lex.europa.eu/eli/reg/2016/679/oj)  
  Supports: Art. 32 security of processing; Art. 33 breach notification, used in the impact scale ('reportable GDPR breach').  
  Used in: Scoring Method; R03, R10, R18. Access: Free.

## Threat context

- [zm-online: Justus-Liebig-Universität ist wieder online (7 Jan 2020)](https://www.zm-online.de/news/detail/justus-liebig-universitaet-ist-wieder-online)  
  Supports: Malware attack (likely Emotet) shut down Uni Gießen in December 2019.  
  Used in: R01. Access: Free web page.
- [The Record: Vice Society claims ransomware attack on HAW Hamburg (6 Mar 2023)](https://therecord.media/germany-ransomware-haw-hamburg-vice-society)  
  Supports: Ransomware at HAW Hamburg (Dec 2022) and University of Duisburg-Essen (Nov 2022).  
  Used in: R01. Access: Free web page.

## Files in this folder

- `saarland-university-risk-register.xlsx`: the register (sheets Read Me, Risk Register, Summary, Scoring Method, Lists, Sources).
- `build_register.py`: Python script that generates the workbook and this file (`python3 build_register.py out.xlsx sources.md`).
- `sources.md`: this list.

The ISO standards are paid and cannot be shared. The BSI and GDPR documents are free to download from the links above.
