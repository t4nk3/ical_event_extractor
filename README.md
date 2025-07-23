# iCal Event Extractor

A desktop GUI app to import XLSM or PDF files containing project schedules, extract time-based data (delivery times, work days, export/send-back times), and export them as iCal (.ics) files.

## Features
- Import XLSM or PDF files
- Extracts events (all-day) with date, event type, project/code, and notes
- Preview and edit events/notes before export
- Export to .ics (iCal) format

## Usage
1. Launch the app
2. Click 'Import' to select an XLSM or PDF file
3. Review and edit extracted events/notes
4. Click 'Export' to save as .ics

## Requirements
- Python 3.8+
- openpyxl
- pdfplumber
- icalendar
- tkinter (standard with Python)

## Installation
```
pip install openpyxl pdfplumber icalendar
```

## License
MIT 