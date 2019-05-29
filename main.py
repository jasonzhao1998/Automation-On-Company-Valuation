import os
import re
import time
import numpy as np
import pandas as pd

ROUNDING_DIGIT = 4
IS = "Income Statement"
BS = "Balance Sheet"
CF = "Cashflow Statement"

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
    elif how is "zero":
        df.at[row_label, df.columns[-yrs:]] = 6
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
    def compare(pair):
        if type(pair[0]) is str:
            return len(pair[0])
        return 0
    return max(sorted(score_dict.items(), key=compare), key=lambda pair: pair[1])[0]


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


def process_is(is_df, growth_rates, yrs_to_predict):
    first_yr = is_df.columns[0]
    last_given_yr = is_df.columns[-1]
    sales = searched_label(is_df.index, "total sales")
    cogs = searched_label(is_df.index, "cost of goods sold")
    gross_income = searched_label(is_df.index, "gross income")
    sgna = searched_label(is_df.index, "sg&a expense")
    other_expense = searched_label(is_df.index, "other expense")
    ebit = searched_label(is_df.index, "ebit")
    nonoperating_income = searched_label(is_df.index, "nonoperating income net")
    interest_expense = searched_label(is_df.index, "interest expense")
    unusual = searched_label(is_df.index, "unusual expense")
    income_tax = searched_label(is_df.index, "income taxes")

    # Insert pretax income row before income taxes
    pretax_df = pd.DataFrame(
        {
            col: '={}+{}-{}-{}'.format(
                excel_cell(is_df, ebit, col), excel_cell(is_df, nonoperating_income, col),
                excel_cell(is_df, interest_expense, col), excel_cell(is_df, unusual, col)
            ) for col in is_df.columns
        }, index=["Pretax Income"]
    )
    is_df = insert_before(is_df, pretax_df, "income taxes")
    pretax = searched_label(is_df.index, "pretax income")

    # Add driver rows to income statement
    add_empty_row(is_df)
    is_df.loc["Driver Ratios"] = np.nan
    is_df.loc["Sales Growth %"] = [np.nan] + [
        '={}/{}-1'.format(
            excel_cell(is_df, sales, is_df.columns[i + 1]),
            excel_cell(is_df, sales, is_df.columns[i])
        ) for i in range(len(is_df.columns) - 1)
    ]
    sales_growth = searched_label(is_df.index, "sales growth %")
    add_empty_row(is_df)
    initialize_ratio_row(is_df, cogs, sales, "COGS Sales Ratio")
    cogs_ratio = searched_label(is_df.index, "cogs sales ratio")
    add_empty_row(is_df)
    initialize_ratio_row(is_df, sgna, sales, "SG&A Sales Ratio")
    sgna_ratio = searched_label(is_df.index, "sg&a sales ratio")
    add_empty_row(is_df)
    initialize_ratio_row(is_df, unusual, ebit, "Unusual Expense EBIT Ratio")
    unusual_ratio = searched_label(is_df.index, "unusual expense ebit ratio")
    add_empty_row(is_df)
    initialize_ratio_row(is_df, income_tax, pretax, "Effective Tax Rate")
    effective_tax = searched_label(is_df.index, "effective tax rate")

    # Add prediction years
    for i in range(yrs_to_predict):
        append_yr_column(is_df)

    # Append growth rates to driver row
    is_df.loc[sales_growth].iloc[-yrs_to_predict:] = growth_rates

    # Calculate driver ratios
    driver_extend(is_df, cogs_ratio, "round", last_given_yr, yrs_to_predict)
    driver_extend(is_df, sgna_ratio, "round", last_given_yr, yrs_to_predict)
    driver_extend(is_df, unusual_ratio, "avg", last_given_yr, yrs_to_predict)
    driver_extend(is_df, effective_tax, "avg", last_given_yr, yrs_to_predict)

    # Calculate fixed variables
    fixed_extend(is_df, nonoperating_income, 'prev', yrs_to_predict)
    fixed_extend(is_df, interest_expense, 'prev', yrs_to_predict)
    fixed_extend(is_df, other_expense, 'prev', yrs_to_predict)

    # Calculate net income
    is_df.loc[searched_label(is_df.index, "net income")] = [
        '={}-{}'.format(excel_cell(is_df, pretax, yr), excel_cell(is_df, income_tax, yr))
        for yr in is_df.columns
    ]

    for i in range(yrs_to_predict):
        cur_yr = is_df.columns[-yrs_to_predict + i]
        prev_yr = is_df.columns[-yrs_to_predict + i - 1]    

        # Calculate variables
        is_df.at[sales, cur_yr] = '={}*(1+{})'.format(
            excel_cell(is_df, sales, prev_yr),
            excel_cell(is_df, sales_growth, cur_yr)
        )
        is_df.at[cogs, cur_yr] = '={}*{}'.format(
            excel_cell(is_df, sales, cur_yr),
            excel_cell(is_df, cogs_ratio, cur_yr)
        )
        is_df.at[gross_income, cur_yr] = '={}-{}'.format(
            excel_cell(is_df, sales, cur_yr),
            excel_cell(is_df, cogs, cur_yr)
        )
        is_df.at[sgna, cur_yr] = '={}*{}'.format(
            excel_cell(is_df, sales, cur_yr),
            excel_cell(is_df, sgna_ratio, cur_yr)
        )
        is_df.at[ebit, cur_yr] = '={}-{}-{}'.format(
            excel_cell(is_df, gross_income, cur_yr),
            excel_cell(is_df, sgna, cur_yr),
            excel_cell(is_df, other_expense, cur_yr)
        )
        is_df.at[unusual, cur_yr] = '={}*{}'.format(
            excel_cell(is_df, ebit, cur_yr),
            excel_cell(is_df, unusual_ratio, cur_yr)
        )
        is_df.at[pretax, cur_yr] = '={}+{}-{}-{}'.format(
            excel_cell(is_df, ebit, cur_yr),
            excel_cell(is_df, nonoperating_income, cur_yr),
            excel_cell(is_df, interest_expense, cur_yr),
            excel_cell(is_df, unusual, cur_yr)
        )
        is_df.at[income_tax, cur_yr] = '={}*{}'.format(
             excel_cell(is_df, pretax, cur_yr),
             excel_cell(is_df, effective_tax, cur_yr)
        )
    return is_df


def process_bs(is_df, bs_df, cf_df, yrs_to_predict):
    first_yr = bs_df.columns[0]
    last_given_yr = bs_df.columns[-1]
    st_receivables = searched_label(bs_df.index, "short term receivable")
    cash_st_investments = searched_label(bs_df.index, "cash short term investment")
    inventories = searched_label(bs_df.index, "inventor")
    other_cur_assets = searched_label(bs_df.index, "other current asset")
    total_cur_assets = searched_label(bs_df.index, "total current asset")
    net_property_plant_equipment = searched_label(bs_df.index, "net property plant qquipment")
    total_investments_n_advances = searched_label(bs_df.index, "total investment advance")
    intangible_assets = searched_label(bs_df.index, "intangible asset")
    deferred_tax_assets = searched_label(bs_df.index, "deferred tax asset")
    other_assets = searched_label(bs_df.index, "other asset")
    total_assets = searched_label(bs_df.index, "total asset")
    st_debt_n_cur_portion_lt_debt = searched_label(bs_df.index, "debt st lt cur portion")
    accounts_payable = searched_label(bs_df.index, "account payable")
    income_tax_payable = searched_label(bs_df.index, "income tax payable")
    other_cur_liabilities = searched_label(bs_df.index, "other current liabilities")
    total_cur_liabilities = searched_label(bs_df.index, "total current liabilities")
    lt_debt = searched_label(bs_df.index, "long term debt")
    provision_for_risks_n_charges = searched_label(bs_df.index, "provision for risk & charge")
    deferred_tax_liabilities = searched_label(bs_df.index, "deferred tax liabilities")
    other_liabilities = searched_label(bs_df.index, "other liabilities")
    total_liabilities = searched_label(bs_df.index, "total liabilities")
    sales = searched_label(is_df.index, "total sales")
    cogs = searched_label(is_df.index, "cost of goods sold")
    net_income = searched_label(is_df.index, "net income")
    deprec_deplet_n_amort = searched_label(cf_df.index, "depreciation depletion amortization") 
    capital_expenditures = searched_label(cf_df.index, "capital expenditure")
    cash_div_paid = searched_label(cf_df.index, "cash div paid")
    charge_in_capital_stock = searched_label(cf_df.index, "charge in capital stock")
    cash_balance = searched_label(cf_df.index, "cash balance")

    # Add working capital row
    add_empty_row(bs_df)
    bs_df.loc["Working Capital"] = [
        '={}-{}'.format(
            excel_cell(bs_df, total_cur_assets, yr),
            excel_cell(bs_df, total_cur_liabilities, yr)
        ) for yr in bs_df.columns
    ]

    # Add driver rows to balance sheet
    add_empty_row(bs_df)
    bs_df.loc["Driver Ratios"] = np.nan
    bs_df.loc["DSO"] = [
        "={}/'{}'!{}*365".format(
            excel_cell(bs_df, st_receivables, yr), IS, excel_cell(is_df, sales, yr)
        ) for yr in bs_df.columns
    ]
    dso = searched_label(bs_df.index, "dso")

    bs_df.loc["Other Current Assets Growth %"] = np.nan
    bs_df.loc["Other Current Assets Growth %"].iloc[1:] = [
        '={}/{}-1'.format(
            excel_cell(bs_df, other_cur_assets, bs_df.columns[i]),
            excel_cell(bs_df, other_cur_assets, bs_df.columns[i + 1])
        ) for i in range(len(bs_df.columns) - 1)
    ]
    other_cur_assets_growth = searched_label(bs_df.index, "other current asset growth %")

    add_empty_row(bs_df)
    bs_df.loc["DPO"] = [
        "={}/'{}'!{}*366".format(
            excel_cell(bs_df, accounts_payable, yr), IS, excel_cell(is_df, cogs, yr)
        ) for yr in bs_df.columns
    ]
    dpo = searched_label(bs_df.index, "dpo")
    bs_df.loc["Miscellaneous Current Liabilities Growth %"] = np.nan
    bs_df.loc["Miscellaneous Current Liabilities Growth %"].iloc[1:] = [
        '={}/{}-1'.format(
            excel_cell(bs_df, other_cur_liabilities, bs_df.columns[i]),
            excel_cell(bs_df, other_cur_liabilities, bs_df.columns[i + 1])
        ) for i in range(len(bs_df.columns) - 1)
    ]
    misc_cur_liabilities_growth = searched_label(bs_df.index, "misc current liabilit growth")
    
    # Add prediction years
    for i in range(yrs_to_predict):
        append_yr_column(bs_df)

    # Calculate driver ratios
    bs_df.loc[dso][last_given_yr:] = '=' + excel_cell(bs_df, dso, bs_df.columns[-yrs_to_predict - 2])
    driver_extend(bs_df, dpo, "avg", last_given_yr, yrs_to_predict)
    driver_extend(bs_df, other_cur_assets_growth, "avg", last_given_yr, yrs_to_predict)
    driver_extend(bs_df, misc_cur_liabilities_growth, "avg", last_given_yr, yrs_to_predict)

    # Calculate fixed variables
    fixed_extend(bs_df, total_investments_n_advances, 'prev', yrs_to_predict)
    fixed_extend(bs_df, intangible_assets, 'prev', yrs_to_predict)
    fixed_extend(bs_df, deferred_tax_assets, 'prev', yrs_to_predict)
    fixed_extend(bs_df, other_assets, 'prev', yrs_to_predict)
    fixed_extend(bs_df, st_debt_n_cur_portion_lt_debt, 'prev', yrs_to_predict)
    fixed_extend(bs_df, income_tax_payable, 'prev', yrs_to_predict)
    fixed_extend(bs_df, lt_debt, 'prev', yrs_to_predict)
    fixed_extend(bs_df, provision_for_risks_n_charges, 'prev', yrs_to_predict)
    fixed_extend(bs_df, deferred_tax_liabilities, 'prev', yrs_to_predict)
    fixed_extend(bs_df, other_liabilities, 'prev', yrs_to_predict)

    for i in range(yrs_to_predict):
        pass
        
    return bs_df


def process_cf(income_statement, balance_sheet, cash_flow, yrs_to_predict):
    row_labels = cash_flow.index
    for i in range(yrs_to_predict):
        append_yr_column(cash_flow)
    fixed_extend(cash_flow, searched_label(row_labels, "deferred taxes"), "zero", yrs_to_predict)
    fixed_extend(cash_flow, searched_label(row_labels, "other funds"), "zero", yrs_to_predict)
    fixed_extend(cash_flow, searched_label(row_labels, "net asset"), "zero", yrs_to_predict)
    fixed_extend(cash_flow, searched_label(row_labels, "fixed asset"), "zero", yrs_to_predict)
    fixed_extend(cash_flow, searched_label(row_labels, "purchase sale investments"),
                 "zero", yrs_to_predict)

    return cash_flow


def initialize_ratio_row(df, top_label, bot_label, new_label):
    """Create a new label and set a fractional formula for initialization."""
    df.loc[new_label] = [
        '={}/{}'.format(excel_cell(df, top_label, col), excel_cell(df, bot_label, col))
        for col in df.columns
    ]


def insert_before(df, new_df, label):
    """Insert new DataFrame before the corresponding label row."""
    index = list(df.index).index(searched_label(df.index, label))
    return pd.concat([df.iloc[:index], new_df, df[index:]])


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

    income_statement = process_is(income_statement, growth_rates, yrs_to_predict)
    balance_sheet = process_bs(income_statement, balance_sheet, cash_flow, yrs_to_predict)
    cash_flow = process_cf(income_statement, balance_sheet, cash_flow, yrs_to_predict)

    print(income_statement)
    print(balance_sheet)

    with pd.ExcelWriter("output.xlsx") as writer:
        income_statement.to_excel(writer, sheet_name=IS)
        balance_sheet.to_excel(writer, sheet_name=BS)
        cash_flow.to_excel(writer, sheet_name=CF)


if __name__ == "__main__":
    main()
