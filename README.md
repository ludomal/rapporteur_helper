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

1. Install required packages
```
pip install -r requirements.txt
```

2. Update meeting information in `generate_reports.py`:

* `meetingDetails`: place and date, example: `"Geneva, 18-26 January 2023"`

*  `meetingDate`: first day of the meeting in the format `YYMMDD`. For example, for the meeting starting January 18, 2023: `"230118"`

2. Execute the script

```
python generate_report.py 
```

Word documents for each question are generated automatically in a directory named as the meeting date, for example `./230118`.
