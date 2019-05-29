import os
import time
import numpy as np
import pandas as pd


ROUNDING_DIGIT = 4

def fixed_extend(df, row_label, how, years):
    if how is "prev":
        df.at[row_label, df.columns[-years:]] = df.loc[row_label, df.columns[-years - 1]]
    elif how is "avg":
        mean = df.loc[row_label].iloc[:-years].mean(axis=0)
        df.at[row_label, df.columns[-years]] = mean
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
    a1 = formula[1]
    n1 = 0
    op = ''
    index = 0
    for i, char in enumerate(formula[2:]):
        if char.isdigit():
            n1 *= 10
            n1 += int(char)
        else:
            op = char
            index = i
            break
    op = formula[index + 2]
    a2 = formula[index + 3]
    n1 -= 2
    n2 = int(formula[index + 4:]) - 2
    col1 = ord(a1) - ord('A') - 1
    col2 = ord(a2) - ord('A') - 1
    if op is '/':
        return df.iat[n1, col1] / df.iat[n2, col2]
    elif op is '*':
        return df.iat[n1, col1] * df.iat[n2, col2]
    elif op is '+':
        return df.iat[n1, col1] + df.iat[n2, col2]
    elif op is '-':
        return df.iat[n1, col1] - df.iat[n2, col2]
    else:
        print("ERROR: Invalid operator symbol")
        exit(1)


def excel_cell(df, row_label, col_label):
    # Note that if there are more than 26 columns, this function does not work
    if not row_label:
        print("ERROR: excel_cell")
        exit(1)
    letter = chr(ord('A') + df.columns.get_loc(col_label) + 1)
    number = str(2 + df.index.get_loc(row_label))
    return letter + number


def searched_label(labels, target):
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

    # Drop rows with elements that are all NaN
    # df = df.dropna(how='all')

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


def append_year_column(df):
    cur_year = str(df.columns[len(df.columns) - 1])
    if cur_year[-1] == 'E':
        cur_year = str(int(cur_year[:-1]) + 1) + 'E'
    else:
        cur_year = str(int(cur_year) + 1) + 'E'
    array = ['0' if i else np.nan for i in df.iloc[:,-1].notna().values]
    df.insert(len(df.columns), cur_year, array)


def append_next_income_statement(income_df, growth_rate, years_to_predict):
    # Uesful short-hands
    row_labels = income_df.index
    last_given_year = income_df.columns[-1]
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

    for i in range(years_to_predict):
        append_year_column(income_df)

    # Append growth rate to driver row
    income_df.loc[sales_growth].iloc[-years_to_predict:] = growth_rate

    # Append rounded COGS ratio to driver row
    cogs_val = round(eval_formula(income_df, income_df.at[cogs_ratio, last_given_year]),
                     ROUNDING_DIGIT)
    income_df.loc[cogs_ratio].iloc[-years_to_predict:] = cogs_val

    # Append rounded SG&A ratio to driver row
    sgna_val = round(eval_formula(income_df, income_df.at[sgna_ratio, last_given_year]),
                     ROUNDING_DIGIT)
    income_df.loc[sgna_ratio].iloc[-years_to_predict:] = sgna_val

    # Append average Unusual Expense ratio to driver row
    unusual_val = np.mean(
        [eval_formula(income_df, income_df.at[unusual_ratio, yr])
        for yr in income_df.columns[:-years_to_predict]]
    )
    income_df.loc[unusual_ratio].iloc[-years_to_predict:] = round(unusual_val,
                                                                  ROUNDING_DIGIT)

    # Calculate fixed variables
    fixed_extend(income_df, searched_label(row_labels, "nonoperating income net"), 'prev',
                 years_to_predict)
    fixed_extend(income_df, searched_label(row_labels, "interest expense"), 'prev',
                 years_to_predict)
    fixed_extend(income_df, searched_label(row_labels, "other expense"), 'prev', years_to_predict)

    for i in range(years_to_predict):
        cur_year = income_df.columns[-years_to_predict + i]
        prev_year = income_df.columns[-years_to_predict + i - 1]    

        # Calculate total sale
        income_df.at[sales, cur_year] = '=' + excel_cell(income_df, sales, prev_year) + \
                                              '*(1+' + \
                                              excel_cell(income_df, sales_growth, cur_year) + \
                                              ')'

        # Calculate COGS
        income_df.at[cogs, cur_year] = '=' + excel_cell(income_df, sales, cur_year) + \
                                             '*' + excel_cell(income_df, cogs_ratio, cur_year)

        # Calculate gross income
        income_df.at[gross_income, cur_year] = '=' + excel_cell(income_df, sales, cur_year) + \
                                                     '-' + excel_cell(income_df, cogs, cur_year)

        # Calculate SG&A expense
        income_df.at[sgna, cur_year] = '=' + excel_cell(income_df, sales, cur_year) + \
                                             '*' + excel_cell(income_df, sgna_ratio, cur_year)

        # Calcualte EBIT
        income_df.at[ebit, cur_year] = '=' + excel_cell(income_df, gross_income, cur_year) + \
                                             '-' + excel_cell(income_df, sgna, cur_year) + \
                                             '-' + excel_cell(
                                                income_df,
                                                searched_label(row_labels, "other expense"),
                                                cur_year
                                            )

        # Calculate unusual expense
        income_df.at[unusual, cur_year] = '=' + excel_cell(income_df, ebit, cur_year) + \
                                                '*' + excel_cell(income_df, unusual_ratio, cur_year)
    return income_df


def append_next_balance_sheet(income_statement, balance_sheet, growth_rate):
    append_year_column(balance_sheet)
    return balance_sheet


def append_next_cash_flow(income_statement, balance_sheet, cash_flow, growth_rate):
    append_year_column(cash_flow)
    return cash_flow


def initialize_ratio_row(df, top_label, bot_label, new_label):
    df.loc[new_label] = [
        '=' + excel_cell(df, searched_label(df.index, top_label), df.columns[i]) + '/' +
        excel_cell(df, searched_label(df.index, bot_label), df.columns[i])
        for i in range(len(df.columns))
    ]


def insert_before(df, data_df, label):
    index = list(df.index).index(searched_label(df.index, label))
    df = pd.concat([df.iloc[:index], data_df, df[index:]])


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
    years_to_predict = len(growth_rates)

    # Add sales growth % driver rows to income statement
    income_statement.loc["null"] = np.nan
    income_statement.index = list(income_statement.index)[:-1] + [np.nan]
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

    # Add COGS % driver rows to income statement
    income_statement.loc["null"] = np.nan
    income_statement.index = list(income_statement.index)[:-1] + [np.nan]
    initialize_ratio_row(income_statement, "cogs", "total sales", "COGS Sales Ratio")

    # Add SG&A % driver row to income statement
    income_statement.loc["null"] = np.nan
    income_statement.index = list(income_statement.index)[:-1] + [np.nan]
    initialize_ratio_row(income_statement, "sg&a expense", "total sales", "SG&A Sales Ratio")

    # Add SG&A % driver row to income statement
    income_statement.loc["null"] = np.nan
    income_statement.index = list(income_statement.index)[:-1] + [np.nan]
    initialize_ratio_row(income_statement, "unusual expense", "ebit", "Unusual Expense EBIT Ratio")

    # Insert pretax income row before income taxes
    

    income_statement = append_next_income_statement(income_statement, growth_rates,
                                                    years_to_predict)
    balance_sheet = append_next_balance_sheet(income_statement, balance_sheet, growth_rates)
    cash_flow = append_next_cash_flow(income_statement, balance_sheet, cash_flow, growth_rates)

    print(income_statement)

    with pd.ExcelWriter("output.xlsx") as writer:
        income_statement.to_excel(writer, sheet_name="Income Statement")
        balance_sheet.to_excel(writer, sheet_name="Balance Sheet")
        cash_flow.to_excel(writer, sheet_name="Cashflow Statement")


if __name__ == "__main__":
    main()
