"""Convert ninja log file to Excel.

Usage:
    ninja_top.py [--open] [--output=OUTPUT_FILE] NINJA_LOG

Arguments:
    NINJA_LOG        Path of the .ninja.log file

Options:
    -h --help               print help
    --open                  open the Excel file when finished (requires Excel installation)
    --output=OUTPUT_FILE    output file. The default is "ninja_log.xlsx" in the same directory as the input file.
"""
from docopt import docopt

import xlsxwriter
import os
import pathlib
import subprocess

def get_ninja_durations(ninja_log_path):
    result = []
    file = open(ninja_log_path, 'r')
    lines = file.readlines()

    # start     end     restat (ignored)    target  command_hash
    # 915240	1175827	616378542	name	a0f6149c34267792
    for line in lines:
        fields = line.split("\t")
        if len(fields) == 5:
            start_time = int(fields[0]) / 1000.0
            end_time = int(fields[1]) / 1000.0
            edge_name = fields[3]
            t = (edge_name, end_time - start_time)
            result.append(t)
    result = sorted(result, key=lambda tup: tup[1], reverse=True)
    return result

if __name__ == '__main__':
    arguments = docopt(__doc__)
    # print(arguments)

    ninja_log_file = arguments["NINJA_LOG"]
    result = get_ninja_durations(ninja_log_file)

    output_file = arguments["--output"]
    if not output_file:
        output_file = str(pathlib.Path(ninja_log_file).parent / "ninja_log.xlsx")
    output_file = pathlib.Path(output_file).absolute()

    print(output_file)

    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    row = 0
    worksheet.write_string(row, 0, "Edge name")
    worksheet.write_string(row, 1, "Duration")
    for t in result:
        row = row + 1
        worksheet.write_string(row, 0, t[0])
        worksheet.write_number(row, 1, t[1])

    # Set width of the first column:
    worksheet.set_column(0, 0, 120)

    workbook.close()

    subprocess.call(["explorer.exe", output_file])
