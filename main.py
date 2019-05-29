import os
import re
import time
import numpy as np
import pandas as pd

# Next things to write:
#   General eval_formula
#   Generalize most indices


ROUNDING_DIGIT = 4


def add_empty_row(df):
    """Adds an empty row to the bottom of DataFrame."""
    df.loc["null"] = np.nan
    df.index = list(df.index)[:-1] + [np.nan]


def driver_extend(df, row_label, how, last_given_yr, yrs_to_predict):
    formula = ""
    if how is "round":
        formula = "=ROUND(" + excel_cell(df, row_label, last_given_yr) + ',' + \
                  str(ROUNDING_DIGIT) + ')'
    elif how is "avg":
        formula = "=AVERAGE(" + excel_cell(df, row_label, df.columns[0]) + ':' + \
                  excel_cell(df, row_label, last_given_yr) + ')'
    df.loc[row_label].iloc[-yrs_to_predict] = formula
    temp = excel_cell(df, row_label, df.columns[-yrs_to_predict])
    df.loc[row_label].iloc[-yrs_to_predict + 1:] = '=' + temp

def fixed_extend(df, row_label, how, yrs):
    """Predicts the corresponding row of data only using data from current row."""
    if how is "prev":
        df.at[row_label, df.columns[-yrs:]] = df.loc[row_label, df.columns[-yrs - 1]]
    elif how is "avg":
        mean = df.loc[row_label].iloc[:-yrs].mean(axis=0)
        df.at[row_label, df.columns[-yrs]] = mean
    elif how is "mix":
        mean = df.loc[row_label].iloc[:-3].mean(axis=0)
        if abs(mean - df.loc[row_label, df.columns[-2]]) > mean * 0.5:
            df.at[row_label, df.columns[-1]] = df.loc[row_label].iloc[:-1].mean(axis=0)
        else:
            df.at[row_label, df.columns[-1]] = df.loc[row_label, df.columns[-2]]
    else:
        print("ERROR: fixed_extend")
        exit(1)

def eval_formula(df, formula):
    """Evaluates an excel formula of a dataframe.
    The mathematical operations must decrease in priority from left to right."""
    ans = 0
    cells = re.findall(r"[A-Z][0-9]*", formula)
    ops = ['+'] + re.findall(r"[+|-|*|/|]", formula)
    
    for i in range(len(cells)):
        row = int(cells[i][1:]) - 2
        col = ord(cells[i][0]) - ord('A') - 1
        if ops[i] is '+':
            ans += df.iat[row, col]
        elif ops[i] is '-':
            ans -= df.iat[row, col]
        elif ops[i] is '*':
            ans *= df.iat[row, col]
        elif ops[i] is '/':
            ans /= df.iat[row, col]
        else:
            print("ERROR: Invalid operator symbol")
            exit(1)
    return ans


def excel_cell(df, row_label, col_label):
    """Returns corresponding excel cell position given row label and column label. 
    Note that if there are more than 26 columns, this function does not work properly."""
    if not row_label:
        print("ERROR: excel_cell")
        exit(1)
    letter = chr(ord('A') + df.columns.get_loc(col_label) + 1)
    number = str(2 + df.index.get_loc(row_label))
    return letter + number


def searched_label(labels, target):
    """Returns target label from a list of DataFrame labels."""
    score_dict = {label: 0 for label in labels}

    for word in target.split():
        for label in labels:
            if word.lower() in str(label).lower():
                score_dict[label] += 1
    # FIXME what to return during a break-even point
    if sum(score_dict.values()) == 0:
        return ""
    return max(score_dict.items(), key=lambda pair: pair[1])[0]


def preprocess(df):
    # Reverse columns
    df = df.loc[:, ::-1]

    # Replace all '-' with 0
    df = df.replace('-', 0)

    # Delete current data
    if df.iat[0, -1] == 'LTM':
        df = df.iloc[:, :-1]

    # Remove the row with number of days
    df = df[1:]

    # Change dates to only years
    df.columns = [
        '20' + ''.join([char for char in column if char.isdigit()]) for column in df.columns
    ]
    return df


def append_yr_column(df):
    """Appends one empty column representing year into DataFrame."""
    cur_yr = str(df.columns[len(df.columns) - 1])
    if cur_yr[-1] == 'E':
        cur_yr = str(int(cur_yr[:-1]) + 1) + 'E'
    else:
        cur_yr = str(int(cur_yr) + 1) + 'E'
    array = ['0' if i else np.nan for i in df.iloc[:,-1].notna().values]
    df.insert(len(df.columns), cur_yr, array)


def append_next_income_statement(income_df, growth_rate, yrs_to_predict):
    # Uesful short-hands
    row_labels = income_df.index
    first_yr = income_df.columns[0]
    last_given_yr = income_df.columns[-1]
    sales = searched_label(row_labels, "total sales")
    sales_growth = searched_label(row_labels, "sales growth %")
    cogs = searched_label(row_labels, "cogs")
    cogs_ratio = searched_label(row_labels, "cogs sales ratio")
    gross_income = searched_label(row_labels, "gross income")
    sgna = searched_label(row_labels, "sg&a expense")
    sgna_ratio = searched_label(row_labels, "sg&a sales ratio")
    ebit = searched_label(row_labels, "ebit")
    unusual = searched_label(row_labels, "unusual expense")
    unusual_ratio = searched_label(row_labels, "unusual ratio")
    pretax = searched_label(row_labels, "pretax income")
    nonoperating_income = searched_label(row_labels, "nonoperating income net")
    interest_expense = searched_label(row_labels, "interest expense")
    other_expense = searched_label(row_labels, "other expense")
    effective_tax = searched_label(row_labels, "effective tax rate")
    income_tax = searched_label(row_labels, "income taxes")

    for i in range(yrs_to_predict):
        append_yr_column(income_df)

    # Append growth rate to driver row
    income_df.loc[sales_growth].iloc[-yrs_to_predict:] = growth_rate

    # Append driver ratios to driver row
    driver_extend(income_df, cogs_ratio, "round", last_given_yr, yrs_to_predict)
    driver_extend(income_df, sgna_ratio, "round", last_given_yr, yrs_to_predict)
    driver_extend(income_df, unusual_ratio, "avg", last_given_yr, yrs_to_predict)
    driver_extend(income_df, effective_tax, "avg", last_given_yr, yrs_to_predict)

    # Calculate fixed variables
    fixed_extend(income_df, nonoperating_income, 'prev', yrs_to_predict)
    fixed_extend(income_df, interest_expense, 'prev', yrs_to_predict)
    fixed_extend(income_df, other_expense, 'prev', yrs_to_predict)

    income_df.loc[searched_label(row_labels, "net income")] = ['=' + excel_cell() + '-' + excel_cell() for i in range(len(income_df.columns))]

    for i in range(yrs_to_predict):
        cur_yr = income_df.columns[-yrs_to_predict + i]
        prev_yr = income_df.columns[-yrs_to_predict + i - 1]    

        # Calculate total sale
        income_df.at[sales, cur_yr] = '=' + excel_cell(income_df, sales, prev_yr) + '*(1+' + \
                                      excel_cell(income_df, sales_growth, cur_yr) + ')'

        # Calculate COGS
        income_df.at[cogs, cur_yr] = '=' + excel_cell(income_df, sales, cur_yr) + \
                                     '*' + excel_cell(income_df, cogs_ratio, cur_yr)

        # Calculate gross income
        income_df.at[gross_income, cur_yr] = '=' + excel_cell(income_df, sales, cur_yr) + \
                                             '-' + excel_cell(income_df, cogs, cur_yr)

        # Calculate SG&A expense
        income_df.at[sgna, cur_yr] = '=' + excel_cell(income_df, sales, cur_yr) + \
                                     '*' + excel_cell(income_df, sgna_ratio, cur_yr)

        # Calcualte EBIT
        income_df.at[ebit, cur_yr] = '=' + excel_cell(income_df, gross_income, cur_yr) + '-' + \
                                     excel_cell(income_df, sgna, cur_yr) + '-' + \
                                     excel_cell(income_df, other_expense, cur_yr)

        # Calculate unusual expense
        income_df.at[unusual, cur_yr] = '=' + excel_cell(income_df, ebit, cur_yr) + \
                                        '*' + excel_cell(income_df, unusual_ratio, cur_yr)

        # Calculate pretax income
        income_df.at[pretax, cur_yr] = '=' + excel_cell(income_df, ebit, cur_yr) + '+' + \
                                       excel_cell(income_df, nonoperating_income, cur_yr) + \
                                       '-' + excel_cell(income_df, interest_expense, cur_yr) + \
                                       '-' + excel_cell(income_df, unusual, cur_yr)

        # Calculate income taxes
        income_df.at[income_tax, cur_yr] = '=' + excel_cell(income_df, pretax, cur_yr) + '*' + \
                                           excel_cell(income_df, effective_tax, cur_yr)

    return income_df


def append_next_balance_sheet(income_statement, balance_sheet, growth_rate):
    append_yr_column(balance_sheet)
    return balance_sheet


def append_next_cash_flow(income_statement, balance_sheet, cash_flow, growth_rate):
    append_yr_column(cash_flow)
    return cash_flow


def initialize_ratio_row(df, top_label, bot_label, new_label):
    df.loc[new_label] = [
        '=' + excel_cell(df, searched_label(df.index, top_label), col) + '/' +
        excel_cell(df, searched_label(df.index, bot_label), col)
        for col in df.columns
    ]


def insert_before(df, data_df, label):
    index = list(df.index).index(searched_label(df.index, label))
    return pd.concat([df.iloc[:index], data_df, df[index:]])


def main():
    income_statement = pd.read_excel("asset/NVIDIA Income Statement.xlsx", header=4, index_col=0)
    balance_sheet = pd.read_excel("asset/NVIDIA Balance Sheet.xlsx", header=4, index_col=0)
    cash_flow = pd.read_excel("asset/NVIDIA Cash Flow.xlsx", header=4, index_col=0)

    income_statement = preprocess(income_statement)
    balance_sheet = preprocess(balance_sheet)
    cash_flow = preprocess(cash_flow)

    # FIXME temporary slices of data
    income_statement = income_statement[:14]
    balance_sheet = balance_sheet[:31]
    cash_flow = cash_flow[:26]

    # FIXME temporary parameters
    growth_rates = [0.5, 0.5, 0.5, 0.5, 0.5]
    yrs_to_predict = len(growth_rates)

    # Add sales growth % driver rows to income statement
    add_empty_row(income_statement)
    income_statement.loc["Driver Ratios"] = np.nan
    income_statement.loc["Sales Growth %"] = [np.nan] + [
        '=' + excel_cell(
            income_statement, searched_label(income_statement.index, "total sales"),
            income_statement.columns[i + 1]
        ) + '/' + excel_cell(
            income_statement, searched_label(income_statement.index, "total sales"),
            income_statement.columns[i]
        ) + '-1' for i in range(len(income_statement.columns) - 1)
    ]

    # Insert pretax income row before income taxes
    ebit = searched_label(income_statement.index, "ebit")
    nonoperating_income = searched_label(income_statement.index, "nonoperating income")
    interest_expense = searched_label(income_statement.index, "interest expense")
    unusual_expense = searched_label(income_statement.index, "unusual expense")
    pretax = pd.DataFrame(
        {
            col: '=' + excel_cell(income_statement, ebit, col) + '+' + \
            excel_cell(income_statement, nonoperating_income, col) + '-' + \
            excel_cell(income_statement, interest_expense, col) + '-' + \
            excel_cell(income_statement, unusual_expense, col)
            for col in income_statement.columns
        }, index=["Pretax Income"]
    )
    income_statement = insert_before(income_statement, pretax, "income taxes")

    # Add driver rows to income statement
    add_empty_row(income_statement)
    initialize_ratio_row(income_statement, "cogs", "total sales", "COGS Sales Ratio")
    add_empty_row(income_statement)
    initialize_ratio_row(income_statement, "sg&a expense", "total sales", "SG&A Sales Ratio")
    add_empty_row(income_statement)
    initialize_ratio_row(income_statement, "unusual expense", "ebit", "Unusual Expense EBIT Ratio")
    add_empty_row(income_statement)
    initialize_ratio_row(income_statement, "income taxes", "pretax", "Effective Tax Rate")

    income_statement = append_next_income_statement(income_statement, growth_rates, yrs_to_predict)
    balance_sheet = append_next_balance_sheet(income_statement, balance_sheet, growth_rates)
    cash_flow = append_next_cash_flow(income_statement, balance_sheet, cash_flow, growth_rates)

    print(income_statement)

    with pd.ExcelWriter("output.xlsx") as writer:
        income_statement.to_excel(writer, sheet_name="Income Statement")
        balance_sheet.to_excel(writer, sheet_name="Balance Sheet")
        cash_flow.to_excel(writer, sheet_name="Cashflow Statement")


if __name__ == "__main__":
    main()
