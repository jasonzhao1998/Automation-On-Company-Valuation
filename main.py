import os
import numpy as np
import pandas as pd


def main():
    cash_flow = pd.read_excel('asset/NVIDIA Cash Flow.xlsx', header=4, index_col=0)
    income_statement = pd.read_excel('asset/NVIDIA Income Statement.xlsx', header=4, index_col=0)
    balance_sheet = pd.read_excel('asset/NVIDIA Balance Sheet.xlsx', header=4, index_col=0)
    print(cash_flow)
    print(income_statement)
    print(balance_sheet)

if __name__ == "__main__":
    main()
