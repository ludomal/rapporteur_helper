# ITU-T Rapporteur's status report generator

This project generates pre-populated status report for Rapporteurs in Study Group 12.

The scripts use `template.docx` as the template generating formatted word documents.
Information fetched from the ITU-T website are:

- Question title
- (co/associate) rapporteur(s) details
- List of contributions
- List of TDs
- Work programme

## How to use

1. Set `meetingDate` in `generate_reports.py` to the desired meeting. The format is `YYMMDD`. For example, for the meeting starting January 18, 2023:
```
  meetingDate = "230118"
```
  
2. Execute the script

```
python generate_report.py 
```

Word documents for each question are generated automatically in a directory named as the meeting date, for example `./230118`.
