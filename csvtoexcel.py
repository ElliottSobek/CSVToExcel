#!/usr/bin/python3

#     CSV To Excel; Converts a CSV file to an Excel file
#     Copyright (C) 2018  Elliott Sobek
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

import os
import sys
import csv

from os.path import basename, getsize, exists, splitext, dirname
from xlsxwriter.workbook import Workbook


def gen_outfile(filepath, extension):
    out_dir = dirname(filepath)
    out_base = basename(filepath)
    base_outfile = splitext(out_base)[0]

    if out_dir:
        if os.name == "nt":
            return out_dir + '\\' + base_outfile + extension
        return out_dir + '/' + base_outfile + extension
    return out_dir + base_outfile + extension


def main(argc=len(sys.argv), argv=sys.argv):
    if argc < 3:
        print("Usage: " + basename(argv[0]) + " <filename.csv ...> <filename>")
        sys.exit(1)

    for i in range(1, argc - 1):
        file = argv[i]
        if not exists(file):
            print("Error: " + file + " does not exist")
            sys.exit(1)
        elif not getsize(file):
            print("Error: " + file + " contains no data/is empty")
            sys.exit(1)
        elif not file.endswith('.csv'):
            print("Error: " + file + " is not comma separated value (csv) format")
            sys.exit(1)

    outfile = gen_outfile(argv[-1], ".xlsx")
    workbook = Workbook(outfile, {'constant_memory': True,
                                  'strings_to_numbers': True})
    for i in range(1, argc - 1):
        csvfile = argv[i]
        worksheet = workbook.add_worksheet(basename(csvfile))

        with open(csvfile, 'r', encoding='utf-8', newline='') as xlsxfile:
            reader = csv.reader(xlsxfile)

            for r, row in enumerate(reader):
                for c, col in enumerate(row):
                    worksheet.write(r, c, col)
    workbook.close()
    return


main()
