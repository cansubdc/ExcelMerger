#!/usr/bin/env python
import os
import sys
import pandas as pd


def merge(files):
    tmp_data = pd.DataFrame()
    for file in files:
        df_file = pd.read_excel(file, index_col=0)
        if not df_file.keys().equals(tmp_data.keys()) and not tmp_data.empty:
            tmp_data.drop(tmp_data.index, inplace=True)
            break
        tmp_data = pd.concat([tmp_data, df_file])
    return tmp_data


if __name__ == '__main__':
    print(os.getcwd())
    print(sys.argv)

    if len(sys.argv) < 2:
        Exception("Gerekli parametreleri girmediniz.")
    try:
        all_data = merge(sys.argv[1:])
        writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
        all_data.to_excel(writer, sheet_name='Sheet1')
        writer.save()
    except FileNotFoundError:
        print("Boyle bir dosya bulunamadi")
