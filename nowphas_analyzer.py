#!/usr/bin/env python3
"""
nowphas_analyzer
"""

__author__ = "Author"
__version__ = "0.1.0"
__license__ = "MIT"

import os
import io
import glob
from datetime import datetime
import argparse
import pandas as pd
import openpyxl as px
from openpyxl.chart import Reference, RadarChart


def read_file(filepath):
    """
    Read NOWPHAS data file into data frame
    """
    columns = ["年月日時", "ﾌﾗｸﾞ", "波数", "平均波波高", "平均波周期", "有義波波高", "有義波周期",
               "1/10波波高", "1/10波周期", "最高波波高", "最高波周期", "波向"]
    offset = 12 if "e" in os.path.basename(filepath) else 10
    buffer = ""
    with open(filepath, mode='r', encoding='shift-jis') as f:
        for line in f:
            buffer += line[:offset].replace(' ', '0') + line[offset:]
    df = pd.read_csv(io.StringIO(buffer),
                     delim_whitespace=True,
                     header=None,
                     skiprows=1,
                     names=columns)
    return df


def read_dir(dirpath):
    """
    All Nowphas data files in a folder are read into a data frame,
    concatenated, and reordered by datetime.
    """
    filepaths = glob.glob(os.path.join(dirpath, "[hH]*.txt"))
    df = pd.concat([read_file(filepath) for filepath in filepaths],
                   ignore_index=True)
    df = df.sort_values("年月日時")
    df = df.reset_index(drop=True)
    return df


def map_datetime(_datetime):
    """
    Convert year, month, date, and time to datetime type
    """
    if len(str(_datetime)) == 10:
        return datetime.strptime(str(_datetime), "%Y%m%d%H")
    else:
        return datetime.strptime(str(_datetime), "%Y%m%d%H%M")


def add_rader_chart(filepath, sheetname):
    """
    Create a radar chart and write it in Excel
    """
    wb = px.load_workbook(filepath)
    ws = wb[sheetname]

    # format
    from openpyxl.styles.numbers import builtin_format_code
    for index in range(3, 15):
        ws.cell(index, 6).number_format = builtin_format_code(10)
        ws.cell(index, 6).number_format = '0.0%'

    # chart
    rmin = 2
    rmax = 14
    cmin = 5
    cmax = 6

    chart = RadarChart()
    chart.type = 'standard'  # 'filled', 'marker', 'standard'
    labels = Reference(ws, min_col=cmin, min_row=rmin+1, max_row=rmax)
    data = Reference(ws, min_col=cmin+1, max_col=cmax,
                     min_row=rmin, max_row=rmax)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.legend.legendPos = "t"
    chart.title = 'Wave Direction'
    chart.height = 10
    chart.width = 10
    chart.anchor = 'B17'

    ws.add_chart(chart)
    wb.save(filepath)


def nowphas_analyzer(dirpath: str) -> None:
    """
    """
    # Read
    df = read_dir(dirpath)
    # convert
    df = df.replace(
        [66.66, 666.6, 6666, 77.77, 777.7, 7777, 99.99, 999.9, 9999], None)
    df = df.dropna()
    df["年月日時"] = df["年月日時"].apply(map_datetime)
    # binning
    bins = [0, 15, 45, 75, 105, 135, 165, 195, 225, 255, 285, 315, 345, 360]
    labels = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 1]
    df['Dir No'] = pd.cut(df["波向"], bins=bins, labels=labels, ordered=False)
    # totaling
    total_df = pd.DataFrame({'num': df.groupby('Dir No')["波数"].sum()})
    total = total_df["num"].sum()
    total_df.insert(0, "prob", total_df["num"].apply(lambda x: x / total))
    total_df.insert(0, "value", list(range(0, 360, 30)))
    total_df.insert(0, "range2", list(range(15, 360, 30)))
    total_df.insert(0, "range1", [i % 360 for i in range(345, 690, 30)])
    # write excel
    filename = "nowphas_wave_frequency_distribution.xlsx"
    sheetname = "sheet1"
    with pd.ExcelWriter(filename) as writer:
        total_df.to_excel(writer, sheetname, startrow=1, startcol=1)
    add_rader_chart(filename, sheetname)


def main(args):
    """ Main entry point of the app """
    nowphas_analyzer(args.dirpath)


if __name__ == "__main__":
    """ This is executed when run from the command line """
    parser = argparse.ArgumentParser()

    # Required positional argument
    parser.add_argument(
        "dirpath",
        help="Path of the directory containing NOWPHAS data files.")

    args = parser.parse_args()
    main(args)
