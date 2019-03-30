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

slidefiletypes = ['.svs', '.ndpi']

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
    if anonid_number <= previous:
        err(row, "AnonID number is {!r}, but must be greater than the previous ({}).",
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

def get_barcode(anonid, block, stain):
    return "{0};{0};{1};{2}".format(anonid, block, stain)

def files(directory, *ext):
    if ext:
        return [f for f in next(os.walk(directory))[2] if os.path.splitext(f)[1] in ext]
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
        slidefiles = files(os.path.join(persondir, slidedir), *slidefiletypes)
        if len(slidefiles) != 1:
            err(i, "Expected 1 slide file in slide directory {!r} for Person {!r} but found {}{}.",
                slidedir, os.path.basename(persondir), len(slidefiles),
                ' ({})'.format(', '.join(repr(f) for f in slidefiles) if slidefiles else ''))
        origfile = os.path.join(persondir, slidedir, slidefiles[0])
        slides.append((block, stain, origfile))
    return slides

def anonymize_slide(i, origfile, anondir, barcode):
    anonymize_cmd = ["anonymize_wsi", "-o", anondir, "-bv", barcode, origfile]
    try:
        subprocess.run(anonymize_cmd, check=True)
    except:
        err(i, "anonymize_wsi returned a nonzero exit code for command {}. Aborting.",
            " ".join(repr(arg) for arg in anonymize_cmd))

def mark_red(cell):
    cell.fill = openpyxl.styles.PatternFill(start_color="FFC7CE", fill_type = "solid")
    cell.font = openpyxl.styles.Font(color="FF0000")

def mark_ok(cell):
    cell.style = 'Normal'

def mark_done(worksheet, row, person, origfile):
    p = worksheet.cell(row, 2)
    o = worksheet.cell(row, 3)
    mark_red(p) if person else mark_ok(p)
    mark_red(o) if origfile else mark_ok(o)

def update_spreadsheet(i, ws, person, origfile, anonid, slides):
    ws.insert_rows(i+1, len(slides) - 1)
    for slideindex, (block, stain, origfile) in enumerate(slides):
        barcode = get_barcode(anonid, block, stain)
        ws.cell(i + slideindex, 1, 'Done')
        ws.cell(i + slideindex, 2, person)
        ws.cell(i + slideindex, 3, origfile)
        ws.cell(i + slideindex, 4, anonid)
        ws.cell(i + slideindex, 5, block)
        ws.cell(i + slideindex, 6, stain)
        if slideindex > 0:
            for c in range(7, ws.max_column):
                ws.cell(i + slideindex, c, ws.cell(i, c).value)
        mark_done(ws, i + slideindex, person, origfile)

def get_str(worksheet, row, column):
    return str(worksheet.cell(row, column).value or '').strip()

def get_personid(person):
    if person.endswith('.zip'):
        return person[:-4]
    return person

def anonymize(wb, basedir, tmpdir, anondir, excelfile):
    ws = wb.active
    check_worksheet(ws)
    prefix = get_str(ws, 9, 5)
    digits = ws["E10"].value or 3
    barcodes = set()
    anonids = dict() # personid -> anonid
    personids = dict() # anonid -> personid

    personid_rows = dict() # personid -> rownumber
    done = set()
    for i in range(15, ws.max_row + 1):
        status = get_str(ws, i, 1)
        personid = get_personid(get_str(ws, i, 2))
        if personid:
            if personid in personid_rows and not (status.lower() == 'done' and personid in done):
                err(i, "Persons must be unique, but ID {!r} occurs multiple times; here and also on row {}.",
                    personid, personid_rows[personid])
            if status.lower() == 'done':
                done.add(personid)
            personid_rows[personid] = i

    # Skip headers
    i = 15
    anonid_number = 0
    person = None
    while i <= ws.max_row:
        # Status	Person	OrigFile	AnonID	Block	Stain
        status = get_str(ws, i, 1)
        person = get_str(ws, i, 2) or person
        origfile = get_str(ws, i, 3)
        anonid = get_str(ws, i, 4)
        block = get_str(ws, i, 5)
        stain = get_str(ws, i, 6)

        if status:
            if status.lower() == "done":
                if not (anonid and block and stain):
                    err(i, "AnonID/Block/Stain missing!")
                mark_done(ws, i, person, origfile)
                wb.save(excelfile)
            elif status.lower() != "ignore":
                err(i, "Unknown Status {!r}.", status)
            i += 1
            continue

        person = os.path.basename(person)
        if not person:
            err(i, "No Person specified.")
        personfile = os.path.join(basedir, person)
        if not os.path.exists(personfile):
            err(i, "Person {!r} does not exist in work directory {}.", person, basedir)
        is_zip = person.endswith('.zip')

        # Validate IDs, mappings etc.
        personid = get_personid(person)
        if anonid:
            anonid_number = get_anonid_number(i, anonid, prefix, digits, anonid_number)
        else:
            anonid, anonid_number = make_anonid(personid, anonids, prefix, digits, anonid_number)
        verify_id_mapping(i, personid, anonid, personids, anonids)

        if origfile or block or stain:
            err(i, "Garbage in OrigFile/Block/Stain columns; OrigFile, Block and Stain must be given by subfolders to Person.")

        persondir = personfile
        if is_zip:
            persondir = os.path.join(tmpdir, personid)
            print("Decompressing {}...".format(person))
            zipfile.ZipFile(personfile).extractall(tmpdir)

        slides = get_slides(i, persondir)
        if not slides:
            err(i, "No slides present for Person {!r}.", person)
        for block, stain, origfile in slides:
            barcode = get_barcode(anonid, block, stain)
            anonymize_slide(i, origfile, anondir, barcode)
            barcodes.add(barcode)

        update_spreadsheet(i, ws, person, origfile, anonid, slides)
        wb.save(excelfile)
        i += len(slides)

        if is_zip:
            shutil.rmtree(persondir)

    return barcodes

def get_garbage(anondir, barcodes):
    garbage = []
    for filename in files(anondir):
        barcode = filename.rsplit('_', 1)[0]
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
    print("\nDone! {} anonymized images added to {!r}. {!r} has been updated.".format(
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
