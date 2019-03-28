# AIDA Pathology Anonymizer for Excel Sheets
Save anonymized copies of Whole Slide Imaging (WSI) pathology data for research
for easy import into a Picture Archice and Communication System (PACS). Generates
anonymous barcodes including information on request, block and stain.

**Research software not approved for clinical use.**

This tools helps reduce monotonous, repetitive and error-prone manual work, letting
the user carry out anonymization and linking by saving original image files in a given
location and noting down study parameters in an Excel spreadsheet. It is possible
to set up this tool to [run with a single click](#setup-to-run-with-single-click),
anonymizing data, assigning anonymous identifiers, detecting common manual errors,
assembling information and making ready for anonymous import.

## Requirements

* anonymize_wsi
* Python3 with packages:
  * openpyxl

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
| Done | P456456456 | 567.svs | P-005 |  | HE | Y | 5 | low |
| Done | P456456456 | 678.svs | P-005 |  | HE2 | Y | 5 | low |
| Done | P456456456 | 789.svs | P-005 |  | HE3 | Y | 5 | low |

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
