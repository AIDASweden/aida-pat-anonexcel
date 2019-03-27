#!/usr/bin/env python3
"""AIDA Pathology Anonymizer for Excel Sheets v%s

Research software not approved for clinical use.

Requirements:
* anonymize_wsi.exe
* Python 3 with the following packages:
    * openpyxl
"""
__version__ = "1.0.0"
__doc__ %= __version__

import argparse
import os
import shutil
import subprocess
import sys
import zipfile

import openpyxl

options = {
    '--version': dict(
        action='version',
        version='%(prog)s ' + __version__),
    'excelfile': dict(
        type=argparse.FileType('rb+'),
        metavar="EXCELFILE.xslx",
        help="An AIDA Pathology Anonymization Excel Sheet. Will be updated in place"),
    '--tmpdir': dict(
        metavar="DIR",
        help="Where to keep temporary files. Default: EXCELFILE_tmp."),
    '--anondir': dict(
        metavar="DIR",
        help="Where to save anonymized slides. Default: EXCELFILE_tmp."),
    ('--suppress-exitcode','-z'): dict(
        action='store_true',
        help="Do not use an exitcode when exiting. Only useful with py -i."),
}
def get_options(argv):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    for name, params in options.items():
        if isinstance(name, str):
            parser.add_argument(name, **params)
        else:
            parser.add_argument(*name, **params)
    return parser.parse_args(argv[1:])

tests = dict(
    D1="AIDA Pathology Anonymization Sheet",
    D9="Prefix:",
    D10="Digits:",
    A14="Status",
    B14="Person",
    C14="OrigFile",
    D14="AnonID",
    E14="Block",
    F14="Stain",
)
def check_worksheet(ws):
    for cell, value in tests.items():
        if ws[cell].value != value:
            row = int(cell[1:])
            err(row, "Unsupported excelfile format. Cell {!r} should be {!r} but is {!r}.",
                cell, value, ws[cell].value)

class ParseError(ValueError):
    """Errors encountered while parsing/anonymizing data from Excel spreadsheet."""

def err(row, msg, *args, **kw):
    raise ParseError(row, msg.format(*args, **kw))

def get_anonid_number(row, anonid, prefix, digits, previous):
    try:
        anonid_number = int(anonid[len(prefix):])
    except:
        anonid_number = -1
    if not (anonid.startswith(prefix) and
            anonid_number > 0 and
            len(anonid) == len(prefix) + digits):
            anonid_number = -1
    if anonid_number < 0:
        err(row, "AnonID is {!r}, but must start with Prefix {!r} followed by a {}-digit number > 0.",
            anonid, prefix, digits)
    if anonid_number < previous:
        err(row, "AnonID number is {!r}, but must be equal or greater than the previous ({}).",
            anonid_number, previous)
    return anonid_number

def make_anonid(personid, anonids, prefix, digits, anonid_number):
    anonid = anonids.get(personid)
    if not anonid:
        anonid_number += 1
        anonid = prefix + str(anonid_number).zfill(digits)
    return anonid, anonid_number

def verify_id_mapping(row, personid, anonid, personids, anonids):
    existing_personid = personids.setdefault(anonid, personid)
    if existing_personid and existing_personid != personid:
        err(row, "Person {} is given AnonID {} on this row, but this AnonID was previously given to Person {}.",
            personid, anonid, existing_personid)
    existing_anonid = anonids.setdefault(personid, anonid)
    if existing_anonid and existing_anonid != anonid:
        err(row, "Person {} is given AnonID {} on this row, but this Person was previously given AnonID {}.",
            personid, anonid, existing_anonid)

def make_barcode(anonid, block, stain):
    return "{0};{0};{1};{2}".format(anonid, block, stain)

def files(directory, ext=None):
    if ext:
        return [f for f in next(os.walk(directory))[2] if f.endswith('.svs')]
    return next(os.walk(directory))[2]

def subdirs(directory):
    return next(os.walk(directory))[1]

def get_slides(i, persondir):
    slides = []
    for slidedir in subdirs(persondir):
        if '_' not in slidedir:
            err(i, "Slide directory {!r} for Person {!r} is not named on the format BLOCK_STAIN (eg 'A_HE').",
                slidedir, os.path.basename(persondir))
        block, stain = slidedir.split('_')
        svsfiles = files(os.path.join(persondir, slidedir), '.svs')
        if len(svsfiles) != 1:
            err(i, "Expected 1 .svs file in slide directory {!r} for Person {!r} but found {}.",
                slidedir, os.path.basename(persondir), len(svsfiles))
        origfile = os.path.join(persondir, slidedir, svsfiles[0])
        slides.append((block, stain, origfile))
    return slides

def anonymize_slide(i, processed_rows, anonid, block, stain, origfile, anondir):
    barcode = make_barcode(anonid, block, stain)
    if barcode in processed_rows:
        err(i, "Duplicate found: Row {} and {} both give slides with AnonID:{!r}, Block:{!r}, Stain:{!r}. Please update and rerun.",
            processed_rows[barcode], i, anonid, block, stain)
    anonfile = barcode + '_anon.svs'
    dst = os.path.join(anondir, anonfile)
    if os.path.exists(dst):
        os.remove(dst)
    anonymize_cmd = ["anonymize_wsi.exe", "-bv", barcode, origfile]
    try:
        subprocess.run(anonymize_cmd, check=True)
    except:
        err(i, "anonymize_wsi.exe returned a nonzero exit code for command {}. Aborting.",
            " ".join(repr(arg) for arg in anonymize_cmd))
    src = os.path.join(os.path.dirname(origfile), anonfile)
    os.rename(src, dst)

def update_spreadsheet(i, ws, anonid, slides, processed_rows):
    ws.insert_rows(i+1, len(slides) - 1)
    for slideindex, (block, stain, origfile) in enumerate(slides):
        barcode = make_barcode(anonid, block, stain)
        processed_rows[barcode] = i + slideindex
        ws.cell(i + slideindex, 1, 'Done')
        ws.cell(i + slideindex, 3, origfile)
        ws.cell(i + slideindex, 4, anonid)
        ws.cell(i + slideindex, 5, block)
        ws.cell(i + slideindex, 6, stain)
        if slideindex > 0:
            for c in range(7, ws.max_column):
                ws.cell(i + slideindex, c, ws.cell(i, c).value)
        mark_red(ws.cell(i + slideindex, 2))
        mark_red(ws.cell(i + slideindex, 3))

def get_str(worksheet, row, column):
    return str(worksheet.cell(row, column).value or '').strip()

def mark_red(cell):
    cell.fill = openpyxl.styles.PatternFill(start_color="FFC7CE", fill_type = "solid")
    cell.font = openpyxl.styles.Font(color="FF0000")

def anonymize(wb, basedir, tmpdir, anondir, excelfile):
    ws = wb.active
    check_worksheet(ws)
    prefix = get_str(ws, 9, 5)
    digits = ws["E10"].value
    processed_rows = {} # barcode -> rownumber
    anonids = dict() # personid -> anonid
    personids = dict() # anonid -> personid

    # Skip headers
    i = 15
    anonid_number = 0
    person = None
    while i <= ws.max_row:
        # Status	ZipFile	AnonFile	AnonID	Block	Stain
        status = get_str(ws, i, 1)
        person = get_str(ws, i, 2) or person
        origfile = get_str(ws, i, 3)
        anonid = get_str(ws, i, 4)
        block = get_str(ws, i, 5)
        stain = get_str(ws, i, 6)

        if not person:
            err(i, "No Person specified.")

        # ID mapping
        personid = person
        is_zip = personid.endswith('.zip')
        if is_zip:
            personid = personid[:-4]
        if anonid:
            anonid_number = get_anonid_number(i, anonid, prefix, digits, anonid_number)
        else:
            anonid, anonid_number = make_anonid(personid, anonids, prefix, digits, anonid_number)
        verify_id_mapping(i, personid, anonid, personids, anonids)

        if status:
            if status.lower() == 'done':
                barcode = make_barcode(anonid, block, stain)
                if barcode in processed_rows:
                    err(i, "Duplicate found: Row {} and {} both give slides with AnonID:{!r}, Block:{!r}, Stain:{!r}. Please update and rerun.",
                        processed_rows[barcode], i, anonid, block, stain)
                processed_rows[barcode] = i
            elif status.lower() != "ignore":
                err(i, "Unknown status {!r}.", status)
            mark_red(ws.cell(i, 2))
            mark_red(ws.cell(i, 3))
            wb.save(excelfile)
            i += 1
            continue

        personfile = os.path.join(basedir, person)
        if not os.path.exists(personfile):
            err(i, "Person {!r} does not exist in work directory {}.", person, basedir)

        persondir = personfile
        if is_zip:
            persondir = os.path.join(tmpdir, personid)
            print("Decompressing {}...".format(person))
            zipfile.ZipFile(personfile).extractall(tmpdir)

        slides = get_slides(i, persondir)
        if not slides:
            err(i, "No slides present for Person {!r}.", person)
        for block, stain, origfile in slides:
            anonymize_slide(i, processed_rows, anonid, block, stain, origfile, anondir)

        update_spreadsheet(i, ws, anonid, slides, processed_rows)
        wb.save(excelfile)
        i += len(slides)

        if is_zip:
            shutil.rmtree(persondir)

    return set(processed_rows.keys())

def get_garbage(anondir, barcodes):
    garbage = []
    for filename in files(anondir):
        barcode = filename[:-len("_anon.svs")]
        if barcode not in barcodes:
            garbage.append(filename)
    return garbage

def main(argv):
    options = get_options(argv)
    wb = openpyxl.load_workbook(options.excelfile)
    base, ext = os.path.splitext(options.excelfile.name)
    anonexcelfile = base + '_anon' + ext
    basedir = os.path.dirname(options.excelfile.name)
    if not options.tmpdir:
        options.tmpdir = os.path.join(basedir, base + '_tmp')
    if not os.path.isdir(options.tmpdir):
        os.makedirs(options.tmpdir, 0o755)
    if not options.anondir:
        options.anondir = os.path.join(basedir, base + '_anon')
    if not os.path.isdir(options.anondir):
        os.makedirs(options.anondir, 0o755)
    try:
        barcodes = anonymize(wb, basedir, options.tmpdir, options.anondir, options.excelfile.name)
    except ParseError as err:
        print("Error on row {}: {}".format(*err.args), file=sys.stderr)
        if options.suppress_exitcode:
            return None
        return 1
    garbage = get_garbage(options.anondir, barcodes)
    if garbage:
        print("\nWarning: Possible garbage files found in {!r} folder. You may want to delete these before proceding:".format(options.anondir), file=sys.stderr)
        print("\n".join(repr(s) for s in garbage), file=sys.stderr)
    print("\nDone! {} anonymized images available in {!r}, {!r} has been updated.".format(
            len(barcodes), os.path.basename(options.anondir), os.path.basename(options.excelfile.name)))
    print("\nYour data is now Pseudonymous.\n")
    print("To make your data Anonymous: Delete all keys associating AnonIDs to "
          "Persons, including the Person and OrigFile file cells in {0!r}. "
          "Obviously, none of the study parameters in {0!r} may contain "
          "identifiers either.".format(os.path.basename(options.excelfile.name)))

if __name__ == "__main__":
    exitcode = main(sys.argv)
    if exitcode is not None:
        sys.exit(exitcode)
