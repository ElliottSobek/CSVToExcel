#!/usr/bin/python3

#     CSV To Excel; Converts a CSV file to an Excel file
#     Copyright (C) 2018  Elliott Sobek (elliottsobek@gmail.com)
#
#     This program is free software: you can redistribute it and/or modify
#     it under the terms of the GNU General Public License as published by
#     the Free Software Foundation, either version 3 of the License, or
#     any later version.
#
#     This program is distributed in the hope that it will be useful,
#     but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#     GNU General Public License for more details.
#
#     You should have received a copy of the GNU General Public License
#     along with this program.  If not, see <http://www.gnu.org/licenses/>.

import os
import optparse
import csv

from os.path import dirname, basename, splitext
from sys import argv, stderr
from xlsxwriter.workbook import Workbook


def gen_outfile_abs_path(filepath: str, extension: str) -> str:
    """
    :param filepath: The relative or absolute filepath of the file
    :param extension: The file extension
    :return: The absolute filepath of the file with the extension
    """

    out_dir = dirname(filepath)
    out_base = basename(filepath)
    base_outfile = splitext(out_base)[0]

    if out_dir:
        if os.name == "nt":
            return out_dir + '\\' + base_outfile + extension
        return out_dir + '/' + base_outfile + extension
    return out_dir + base_outfile + extension


def main() -> int:
    parser = optparse.OptionParser(usage="Usage: python3 %prog [options] <filename.csv ...> <outfile>",
                                   version="%prog 1.2")

    parser.add_option("-s", action="store_true", dest="str_to_int_flag",
                      help="Converts strings to integers when writing to excel file")
    parser.add_option("-f", action="store_true", dest="force_flag",
                      help="Forces writing to excel file if one or more csv files are empty")
    parser.add_option("-q", action="store_true", dest="quiet_flag", help="Suppress the startup banner")

    options, args = parser.parse_args()
    exit_failure = 1

    if len(args) < 2:
        print("Usage: " + basename(argv[0]) + "[h] [sfq] [--version] <filename.csv ...> <filename>", file=stderr)
        raise SystemExit(exit_failure)

    in_files = args[:-1]
    outfile = args[-1]

    for csvfile in in_files:
        if not csvfile.endswith(".csv"):
            print("Error: " + csvfile + " is not comma separated value (csv) format", file=stderr)
            raise SystemExit(exit_failure)

    extension = ".xlsx"

    if not outfile.endswith(extension):
        outfile = gen_outfile_abs_path(outfile, extension)

    if not options.quiet_flag:
        print("CSV To Excel (C) 2018  Elliott Sobek (elliottsobek@gmail.com)\n"
              "This program comes with ABSOLUTELY NO WARRANTY.\n"
              "This is free software, and you are welcome to redistribute it under certain conditions.")

    workbook = Workbook(outfile, {"strings_to_numbers": options.str_to_int_flag})

    for csvfile in in_files:
        worksheet = workbook.add_worksheet(basename(csvfile[:-4]))

        try:
            xlsxfile = open(csvfile, mode='r', buffering=1, encoding="utf-8-sig", newline='')
        except (FileNotFoundError, PermissionError) as e:
            print(e, file=stderr)
            try:
                workbook.close()
            except PermissionError as e:
                print(e, file=stderr)
                raise SystemExit(exit_failure)
            raise SystemExit(exit_failure)

        if not os.stat(csvfile).st_size and not options.force_flag:
            print("Error: " + csvfile + " contains no data/is empty", file=stderr)
            raise SystemExit(exit_failure)

        csv_reader = csv.reader(xlsxfile)

        for r, row in enumerate(csv_reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col)
        xlsxfile.close()

    try:
        workbook.close()
    except PermissionError as e:
        print(e, file=stderr)
        raise SystemExit(exit_failure)
    return 0


main()
