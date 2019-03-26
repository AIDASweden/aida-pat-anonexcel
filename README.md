# AIDA Pathology Anonymizer for Excel Sheets
Save anonymized copies of Whole Slide Imaging (WSI) pathology data for research
for easy import into a Picture Archice and Communication System (PACS). Generates
anonymous barcodes including information on request, block and stain.

**Research software not approved for clinical use.**

This tools helps reduce monotonous, repetitive and error-prone manual work, letting
the user work by saving images in the right place and noting study data in an Excel
spreadsheet. This tool can then be set up to
[run with a single click](#setup-to-run-with-single-click), anonymizing data, assigning
anonymous identifiers, detecting common manual errors, assembling information and
making ready for anonymous import.

## Requirements

* anonymize_wsi
* Python3 with packages:
  * openpyxl

## Usage
`py aida-pat-anonexcel.py anonymization-sheet.xlsx`

Make a copy of anonymization-sheet.xlsx and fill it out. Follow instructions in sheet.
Person should be whatever helps you remember who is who, and allows you to store their
images separately in subdirectories with corresponding names.
Running the above command puts anonymous images for export in subdirectory `anon/`
and produces a report in `anonymization-sheet_anon.xlsx` for inspection, adding red
highlights to cells that must be deleted to complete anonymization.

### Example

anonymization-sheet.xlsx:

| Person | AnonID | Block | Stain | Image file | AnonFile | Study parameter 1 | 2 | … |
| --- | --- |   --- | --- |  --- | --- |  --- | --- |  --- |
|P12312124124| | A | HE | orig.svs |  | X | 1 | high |
|            | | B | HE | orig.svs |  | Y | 5 | low |


anonymization-sheet_anon.xlsx:

| Person | AnonID | Block | Stain | Image file | AnonFile | Study parameter 1 | 2 | … |
| --- | --- |   --- | --- |  --- | --- |  --- | --- |  --- |
|P12312124124| MYPROJ-001 | A | HE | orig.svs | MYPROJ-001;A;HE_anon.svs | X | 1 | high |
|            | MYPROJ-001 | B | HE | orig.svs | MYPROJ-001;B;HE_anon.svs | Y | 5 | low |


### Setup to run with single click
You can set up a shortcut to run aida-pat-anonexcel.py on a specific AIDA anonymization
sheet, or to run it on any sheet that you drag-and-drop onto the Shortcut. Both these
methods will open a terminal window for status and error messages, which you can close
when you have finished reading.

1. Install dependencies (Python3 and openpyxl).
2. Put aida-pat-anonexcel.py somewhere permanent where you can find it. Name it myproj.xlsx or similar.
3. Put your copy of the anonymization sheet in a folder where you want to work. Put images in (subfolders to)
this folder.
4. Make a shortcut some place convenient, eg in your work folder or elsewhere.

Make a shortcut:

1. Right click in file explorer > New > Shortcut.
2. Target: Find aida-pat-anonexcel.py.
3. Name it "Anonymize MyProj" or "Drop sheet on me to anonymize" or similar.
4. OK.

Configure your shortcut:

1. Right click your link, choose Properties.
2. Change Target to eg: `py -i "C:\path\to\aida-pat-anonexcel.py" -z "C:\path\to\myproj.xlsx"`
3. Your shortcut is now set up to always run on myproj.xlsx.
4. If you want drag-and-drop instead just delete the last `"C:\path\to\myproj.xlsx"` part.
