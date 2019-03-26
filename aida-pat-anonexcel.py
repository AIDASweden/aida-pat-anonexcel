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

import openpyxl

options = {
    '--version': dict(
        action='version',
        version='%(prog)s ' + __version__),
    'excelfile': dict(
        type=argparse.FileType('rb'),
        metavar="EXCELFILE.xslx",
        help="An AIDA Pathology Anonymization Excel Sheet. Report will be saved to EXCELFILE_anon.xslx."),
    '--anondir': dict(
        metavar="DIR",
        help="Where to save anonymized images. Default: anon/ folder in same folder as EXCELFILE."),
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
    B1="AIDA Pathology Anonymization Sheet",
    B9="Prefix:",
    B10="Digits:",
    A14="Person",
    B14="AnonID",
    C14="Block",
    D14="Stain",
    E14="Image file",
    F14="AnonFile",
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

def get_image_file(row, image_file_name, basedir, personid):
    image_file_name_root, ext = os.path.splitext(image_file_name)
    if ext:
        if ext != ".svs":
            err(row, "Unsupported file name on line {}. Only .svs files are supported.",
                rowindex)
        image_file = image_file_name
    else:
        image_file = image_file_name_root + '.svs'
    sourcedir = basedir
    if not os.path.isfile(os.path.join(sourcedir, image_file)):
        sourcedir = os.path.join(basedir, str(personid))
        image_file = os.path.join(sourcedir, image_file)
        if not os.path.isfile(image_file):
            err(row, "Image file {!r} not found, neither in base directory or Person subdirectory {}/",
                image_file_name, personid)
    return image_file

def process_image_file(row, image_file_name, anonfile, basedir, anondir, personid, barcode):
    image_file = get_image_file(row, image_file_name, basedir, personid)
    anonymize_cmd = ["anonymize_wsi.exe", "-bv", barcode, image_file]
    try:
        subprocess.run(anonymize_cmd, check=True)
    except:
        err(row, "anonymize_wsi.exe returned a nonzero exit code for command {}. Aborting.",
            " ".join(repr(arg) for arg in anonymize_cmd))
    src = os.path.join(os.path.dirname(image_file), anonfile)
    os.rename(src, os.path.join(anondir, anonfile))

def get_str(cell):
    return str((cell.value or '')).strip()

def mark_red(cell):
    cell.fill = openpyxl.styles.PatternFill(start_color="FFC7CE", fill_type = "solid")
    cell.font = openpyxl.styles.Font(color="FF0000")

def anonymize(wb, basedir, anondir, anonexcelfile):
    ws = wb.active
    check_worksheet(ws)
    prefix = ws["C9"].value.strip()
    digits = ws["C10"].value
    rows = iter(ws.rows)
    rowoffset = 0
    processed_rows = dict() # Barcode => Row
    anonids = dict() # ID -> anonid
    personids = dict() # anonid -> ID

    # Skip headers
    for row in rows:
        rowoffset += 1
        if row[0].value == "Person":
            break

    anonid_number = 0
    personid = None
    for i, current_row in enumerate(rows):
        row = rowoffset + i + 1
        # ID  AnonID	Block	Stain	Image file	AnonFile
        personid = get_str(current_row[0]) or personid
        anonid, block, stain, image_file_name, given_anonfile = (get_str(c) for c in current_row[1:6])

        if anonid:
            anonid_number = get_anonid_number(row, anonid, prefix, digits, anonid_number)
        else:
            anonid, anonid_number = make_anonid(personid, anonids, prefix, digits, anonid_number)

        if personid:
            verify_id_mapping(row, personid, anonid, personids, anonids)

        barcode = "{0};{0};{1};{2}".format(anonid, block or '', stain or '')
        if barcode in processed_rows:
            err(row, "Duplicate found: Row {} and {} both have AnonID:{!r}, Block:{!r}, Stain:{!r}. Please update and rerun.",
                processed_rows[barcode], row, anonid, block, stain)

        anonfile = barcode + '_anon.svs'
        if given_anonfile and given_anonfile != anonfile:
            err(row, "AnonFile is {!r} but should be {!r}.", given_anonfile, anonfile)

        anonpath = os.path.join(anondir, anonfile)
        if not (os.path.isfile(anonpath) and os.path.getsize(anonpath)):
            process_image_file(row, image_file_name, anonfile, basedir, anondir, personid, barcode)

        current_row[1].value = anonid
        current_row[5].value = anonfile
        mark_red(current_row[0])
        mark_red(current_row[4])
        wb.save(anonexcelfile)
        processed_rows[barcode] = row
    return set(processed_rows.keys())

def get_garbage(anondir, barcodes):
    garbage = []
    # (dirpath, dirnames, filenames)
    for filename in next(os.walk(anondir))[2]:
        barcode = filename[:-len("_anon.svs")]
        if barcode not in barcodes:
            garbage.append(filename)
    return garbage

def main(argv):
    options = get_options(argv)
    basedir = os.path.dirname(options.excelfile.name)
    anondir = os.path.join(basedir, 'anon')
    if not os.path.isdir(anondir):
        os.makedirs(anondir, 0o755)
    wb = openpyxl.load_workbook(options.excelfile)
    base, ext = os.path.splitext(options.excelfile.name)
    anonexcelfile = base + '_anon' + ext
    try:
        barcodes = anonymize(wb, basedir, anondir, anonexcelfile)
    except ParseError as err:
        print("Error on row {}: {}".format(*err.args), file=sys.stderr)
        if options.suppress_exitcode:
            return None
        return 1
    garbage = get_garbage(anondir, barcodes)
    if garbage:
        print("\nWarning: Possible garbage files found in anon/ folder. You may want to delete these before proceding:", file=sys.stderr)
        print("\n".join(repr(s) for s in garbage), file=sys.stderr)
    print("\nDone! {} anonymized images available in anon/, report in {!r}.".format(
            len(barcodes), os.path.basename(anonexcelfile)))
    print("\nYour data is now Pseudonymous.\n")
    print("To make your data Anonymous: Delete all keys associating AnonIDs to "
          "Persons, including the Person and Image file cells in {0!r} and any "
          "intermediary data like AnonIDs in {1!r}. Obviously, none of the "
          "study parameters in {0!r} may contain identifiers either.".format(
            os.path.basename(anonexcelfile),
            os.path.basename(options.excelfile.name)))

if __name__ == "__main__":
    exitcode = main(sys.argv)
    if exitcode is not None:
        sys.exit(exitcode or 0)
