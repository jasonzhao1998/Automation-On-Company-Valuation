import os
import xlsxwriter
from openpyxl import load_workbook
import numpy as np
import pandas as pd


def test1():
    df = pd.read_excel('pandas_simple.xlsx')
    sum_row = df[["A", "B"]].T.sum()
    print(sum_row)
    df["C"] = sum_row
    print(df)
    df.to_excel("idk.xlsx")


def test():
    df = pd.DataFrame({'A': range(1, 6), 'B': range(10, 0, -2)})
    writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1')
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    worksheet.write_formula('D1', 'D')
    worksheet.write_formula('D2', '=SUM(B2, C2)')
    worksheet.write_formula('D3', '=SUM(B3, C3)')
    writer.save()

    #df.eval('C = A + B', inplace=True)


def main():
    cash_flow = pd.read_excel('asset/NVIDIA Cash Flow.xlsx', header=4, index_col=0)
    income_statement = pd.read_excel('asset/NVIDIA Income Statement.xlsx', header=4, index_col=0)
    balance_sheet = pd.read_excel('asset/NVIDIA Balance Sheet.xlsx', header=4, index_col=0)
    print(cash_flow)
    print(income_statement)
    print(balance_sheet)

if __name__ == "__main__":
    test()
    wb = load_workbook(filename = 'pandas_simple.xlsx')
    sheet_names = wb.get_sheet_names()
    name = sheet_names[0]
    sheet_ranges = wb[name]
    df = pd.DataFrame(sheet_ranges.values)
    print(df)
    # main()
