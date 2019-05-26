import os
import numpy as np
import pandas as pd


def preprocess(df):
    # Reverse columns
    df = df.loc[:, ::-1]

    # Drop rows with elements that are all NaN
    df = df.dropna(how='all')

    # Replace all '-' with 0
    df = df.replace('-', 0)

    # Delete current data
    if df.iat[0, -1] == 'LTM':
        df = df.iloc[:, :-1]

    # Remove rows with label NaN
    df = df[df.index.notna()]

    # Change dates to only years
    df.columns = [
        '20' + ''.join([char for char in column if char.isdigit()]) for column in df.columns
    ]
    return df


def append_year_columns(df, years_to_predict):
    for i in range(years_to_predict):
        cur_year = df.columns[len(df.columns) - 1]
        if cur_year[-1] == 'E':
            cur_year = str(int(cur_year[:-1]) + 1) + 'E'
        else:
            cur_year = str(int(cur_year) + 1) + 'E'
        df.insert(len(df.columns), cur_year, 0)


def append_next_income_statement(income_statement, growth_rates):
    return income_statement


def append_next_balance_sheet(income_statement, balance_sheet, growth_rates):
    return balance_sheet


def append_next_cash_flow(income_statement, balance_sheet, cash_flow, growth_rates):
    return cash_flow


def main():
    income_statement = pd.read_excel('asset/NVIDIA Income Statement.xlsx', header=4, index_col=0)
    balance_sheet = pd.read_excel('asset/NVIDIA Balance Sheet.xlsx', header=4, index_col=0)
    cash_flow = pd.read_excel('asset/NVIDIA Cash Flow.xlsx', header=4, index_col=0)

    income_statement = preprocess(income_statement)
    balance_sheet = preprocess(balance_sheet)
    cash_flow = preprocess(cash_flow)

    # FIXME temporary slices of data
    income_statement = income_statement[:13]
    balance_sheet = balance_sheet[:25]
    cash_flow = cash_flow[:20]

    # FIXME temporary parameters
    growth_rates = []
    years_to_predict = 5

    # Append empty year columns
    append_year_columns(income_statement, years_to_predict)
    append_year_columns(balance_sheet, years_to_predict)
    append_year_columns(cash_flow, years_to_predict)

    for i in range(years_to_predict):
        income_statement = append_next_income_statement(income_statement, growth_rates)
        balance_sheet = append_next_balance_sheet(income_statement, balance_sheet, growth_rates)
        cash_flow = append_next_cash_flow(income_statement, balance_sheet, cash_flow, growth_rates)

    with pd.ExcelWriter('output.xlsx') as writer:
        income_statement.to_excel(writer, sheet_name='Income Statement')
        balance_sheet.to_excel(writer, sheet_name='Balance Sheet')
        cash_flow.to_excel(writer, sheet_name='Cashflow Statement')


if __name__ == "__main__":
    main()
