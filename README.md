# AIDA Pathology Anonymizer for Excel Sheets
A tool for pathology researchers to help reduce repetitive and error-prone
manual work in anonymizing data exports in Whole Slide Imaging (WSI).
This tool lets researchers anonymize their data more easily by saving their data
exports case by case in a folder, and noting them down along with associated
study parameters in an Excel spreadsheet.

From there, this tool can anonymize the slides
[with a single click](#setup-to-run-with-single-click), and check for common
manual errors, mark data red that needs deleting to complete anonymization, and
assemble the anonymized data in a folder ready for easy import into a Picture 
Archice and Communication System (PACS).

**Note: Research software not approved for clinical use.**

## Requirements

* anonymize_wsi
* Python3 with packages:
  * openpyxl

anonymize_wsi is expected to anonymize a single WSI when invoked with:
`anonymize_wsi -o ANONDIR -bv BARCODE FILE`

## Usage
`py aida-pat-anonexcel.py anonymization-sheet.xlsx`

Make a copy of anonymization-sheet.xlsx and fill it out according to instruction in sheet.
Person should be the name of (possibly uncompressed) case export zipfiles with
BLOCK_STAIN subdirectories. Running the above command puts anonymized images for
export in subdirectory `anonymization-sheet_anon/` and updates
`anonymization-sheet_anon.xlsx` to match, with block and stain information on new
rows for each anonymized slide, and with study parameters carried over for each
case and adding red highlight to cells that need be deleted to complete anonymization.

### Example

anonymization-sheet.xlsx (before):

| Status | Person | OrigFile | AnonID | Block | Stain | Study parameter 1 | 2 | … |
| --- | --- | --- | --- | --- | --- |  --- | --- | --- |
| |P123123123| | P-004 |  |  | X | 1 | high |
| |P456456456| | |  |  | Y | 5 | low |

anonymization-sheet.xlsx (after):

| Status | Person | OrigFile | AnonID | Block | Stain | Study parameter 1 | 2 | … |
| --- | --- | --- | --- |  --- | --- | --- | --- | --- |
| Done | P123123123 | 123.svs | P-004 | A | HE | X | 1 | high |
| Done | P123123123 | 234.svs | P-004 | B | HE | X | 1 | high |
| Done | P123123123 | 345.svs | P-004 | C | HE | X | 1 | high |
| Done | P123123123 | 456.svs | P-004 | D | HE | X | 1 | high |
| Done | P456456456 | 567.ndpi | P-005 | A | HE | Y | 5 | low |
| Done | P456456456 | 678.ndpi | P-005 | A | HE2 | Y | 5 | low |
| Done | P456456456 | 789.ndpi | P-005 | A | HE3 | Y | 5 | low |

### Setup to run with single click
You can set up a shortcut to run aida-pat-anonexcel.py either on a specific AIDA anonymization
sheet, or on any sheet that you drag-and-drop onto the shortcut. Both these
methods will open a terminal window for status and error messages which you can close
when you have finished reading.

1. Install dependencies (Python3 and openpyxl).
2. Put aida-pat-anonexcel.py somewhere permanent where you can find it.
3. Put your copy of the anonymization sheet in a folder where you want to work. Name it myproj.xlsx or similar. Put images in (subfolders to) this folder.
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
