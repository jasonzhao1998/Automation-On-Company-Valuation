"""Main program's implementation."""
import openpyxl
import numpy as np
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from style import style_ws
from helper import *

NAME = "NVIDIA"
IS = "Income Statement"
BS = "Balance Sheet"
CF = "Cashflow Statement"

"""
TODO:
    Get name of corporation automatically.
    Customize number of years to consider.
    Million billion problem.
    Optimize searched label.
"""

def preprocess(df):
    """Data cleaning."""
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
    fye = df.columns[1]
    df.columns = [
        '20' + ''.join([char for char in column if char.isdigit()]) for column in df.columns
    ]
    
    # Manage duplicate labels
    ignore = [searched_label(df.index, 'other funds')]
    unique_dict = {}

    duplicate = []
    for idx, label in enumerate(df.index):
        if label in ignore:
            continue
        elif label in unique_dict:
            duplicate.append(idx)
        elif True in list(pd.notna(df.iloc[idx])):
            unique_dict[label] = True
    df = df.iloc[[i for i in range(len(df.index)) if i not in duplicate], :]

    return df, fye


def process_is(is_df, cf_df, growth_rates, yrs_to_predict):
    """Manipulates income statement."""
    # Short-hands
    last_given_yr = is_df.columns[-1]

    # Insert 4 empty rows
    is_df = pd.concat(
        [pd.DataFrame({yr: [np.nan] * 4 for yr in is_df.columns}, index=[np.nan] * 4), is_df]
    )
    cf_df = pd.concat(
        [pd.DataFrame({yr: [np.nan] * 4 for yr in cf_df.columns}, index=[np.nan] * 4), cf_df]
    )

    # Declare income statement labels
    sales = searched_label(is_df.index, "total sales")
    cogs = searched_label(is_df.index, "cost of goods sold COGS including")
    is_df.index = [  # change label name from incl. to excl.
        i if i != cogs else "Cost of Goods Sold (COGS) excl. D&A" for i in list(is_df.index)
    ]
    cogs = "Cost of Goods Sold (COGS) excl. D&A"
    gross_income = searched_label(is_df.index, "gross income")
    sgna = searched_label(is_df.index, "sg&a expense")
    other_expense = searched_label(is_df.index, "other expense")
    ebit = searched_label(is_df.index, "ebit operating income")
    nonoperating_income = searched_label(is_df.index, "nonoperating income net")
    interest_expense = searched_label(is_df.index, "interest expense")
    unusual = searched_label(is_df.index, "unusual expense")
    income_tax = searched_label(is_df.index, "income taxes")
    diluted_eps = searched_label(is_df.index, "eps diluted")
    net_income = searched_label(is_df.index, "net income")
    div_per_share = searched_label(is_df.index, "dividends per share")
    ebitda = searched_label(is_df.index, "ebitda")
    diluted_share_outstanding = "Diluted Shares Outstanding"

    # Drop EBITDA if exists
    if ebitda:
        is_df.drop(ebitda, inplace=True)

    # Insert pretax income row before income taxes
    pretax_df = pd.DataFrame(
        {
            yr: '={}+{}-{}-{}'.format(
                excel_cell(is_df, ebit, yr), excel_cell(is_df, nonoperating_income, yr),
                excel_cell(is_df, interest_expense, yr), excel_cell(is_df, unusual, yr)
            ) for yr in is_df.columns
        }, index=["Pretax Income"]
    )
    is_df = insert_before(is_df, pretax_df, income_tax)
    pretax = "Pretax Income"

    # Insert depreciation & amortization expense before SG&A expense
    dna_expense_df = pd.DataFrame(
        {
            yr: "='{}'!".format(CF) + excel_cell(
                cf_df,
                searched_label(cf_df.index, "depreciation depletion & amortization expense"),
                yr
            ) for yr in is_df.columns
        }, index=["Depreciation & Amortization Expense"]
    )
    is_df = insert_before(is_df, dna_expense_df, sgna)
    dna_expense = "Depreciation & Amortization Expense"

    # Write formulas for COGS excl. D&A
    is_df.loc[cogs] = [
        '={}-{}'.format(is_df.at[cogs, yr], excel_cell(is_df, dna_expense, yr))
        for yr in is_df.columns
    ]

    # Add driver rows to income statement
    add_empty_row(is_df)
    is_df.loc["Driver Ratios"] = np.nan
    add_growth_rate_row(is_df, sales, "Sales Growth %")
    sales_growth = "Sales Growth %"
    add_empty_row(is_df)
    initialize_ratio_row(is_df, cogs, sales, "COGS Sales Ratio")
    cogs_ratio = "COGS Sales Ratio"
    add_empty_row(is_df)
    initialize_ratio_row(is_df, sgna, sales, "SG&A Sales Ratio")
    sgna_ratio = "SG&A Sales Ratio"
    add_empty_row(is_df)
    initialize_ratio_row(is_df, dna_expense, sales, "D&A Sales Ratio")
    dna_ratio = "D&A Sales Ratio"
    add_empty_row(is_df)
    initialize_ratio_row(is_df, unusual, ebit, "Unusual Expense EBIT Ratio")
    unusual_ratio = "Unusual Expense EBIT Ratio"
    add_empty_row(is_df)
    initialize_ratio_row(is_df, income_tax, pretax, "Effective Tax Rate")
    effective_tax = "Effective Tax Rate"

    # Add prediction years
    for i in range(yrs_to_predict):
        add_yr_column(is_df)

    # Append growth rates to driver row
    is_df.loc[sales_growth].iloc[-yrs_to_predict:] = growth_rates

    # Write formulas for EBITDA
    ebitda_df = pd.DataFrame(
        {
            yr: '={}+{}'.format(
                excel_cell(is_df, dna_expense, yr), excel_cell(is_df, ebit, yr)
            ) for yr in is_df.columns
        }, index=[ebitda]
    )
    is_df = insert_after(is_df, ebitda_df, div_per_share)

    # Write formulas for driver ratios
    initialize_ratio_row(is_df, div_per_share, diluted_eps, "Dividend Payout Ratio")
    initialize_ratio_row(is_df, ebitda, sales, "EBITDA Margin", sales_growth)
    is_df.loc[dna_ratio].iloc[-yrs_to_predict:] = is_df.loc[dna_ratio, last_given_yr]
    driver_extend(is_df, cogs_ratio, "round", last_given_yr, yrs_to_predict)
    driver_extend(is_df, sgna_ratio, "round", last_given_yr, yrs_to_predict)
    driver_extend(is_df, unusual_ratio, "avg", last_given_yr, yrs_to_predict)
    driver_extend(is_df, diluted_share_outstanding, "avg", last_given_yr, yrs_to_predict, 3)
    driver_extend(is_df, effective_tax, "avg", last_given_yr, yrs_to_predict)

    # Calculate fixed variables
    fixed_extend(is_df, other_expense, "prev", yrs_to_predict)
    fixed_extend(is_df, nonoperating_income, "prev", yrs_to_predict)
    fixed_extend(is_df, interest_expense, "prev", yrs_to_predict)
    fixed_extend(is_df, div_per_share, "prev", yrs_to_predict)  # FIXME
    fixed_extend(is_df, diluted_share_outstanding, "prev", yrs_to_predict)

    # Write formula for net income
    is_df.loc[net_income] = [
        '={}-{}'.format(excel_cell(is_df, pretax, yr), excel_cell(is_df, income_tax, yr))
        for yr in is_df.columns
    ]

    for i in range(yrs_to_predict):
        cur_yr = is_df.columns[-yrs_to_predict + i]
        prev_yr = is_df.columns[-yrs_to_predict + i - 1]

        # Write formulas
        is_df.at[sales, cur_yr] = '={}*(1+{})'.format(
            excel_cell(is_df, sales, prev_yr), excel_cell(is_df, sales_growth, cur_yr)
        )
        is_df.at[cogs, cur_yr] = '={}*{}'.format(
            excel_cell(is_df, sales, cur_yr), excel_cell(is_df, cogs_ratio, cur_yr)
        )
        is_df.at[gross_income, cur_yr] = '={}-{}'.format(
            excel_cell(is_df, sales, cur_yr), excel_cell(is_df, cogs, cur_yr)
        )
        is_df.at[dna_expense, cur_yr] = '={}*{}'.format(
            excel_cell(is_df, sales, cur_yr), excel_cell(is_df, dna_ratio, cur_yr)
        )
        is_df.at[sgna, cur_yr] = '={}*{}'.format(
            excel_cell(is_df, sales, cur_yr), excel_cell(is_df, sgna_ratio, cur_yr)
        )
        is_df.at[ebit, cur_yr] = '={}-{}'.format(
            excel_cell(is_df, gross_income, cur_yr),
            sum_formula(is_df, ebit, cur_yr, gross_income, 1)
        )
        is_df.at[unusual, cur_yr] = '={}*{}'.format(
            excel_cell(is_df, ebit, cur_yr), excel_cell(is_df, unusual_ratio, cur_yr)
        )
        is_df.at[pretax, cur_yr] = '={}+{}-{}-{}'.format(
            excel_cell(is_df, ebit, cur_yr), excel_cell(is_df, nonoperating_income, cur_yr),
            excel_cell(is_df, interest_expense, cur_yr), excel_cell(is_df, unusual, cur_yr)
        )
        is_df.at[income_tax, cur_yr] = '={}*{}'.format(
             excel_cell(is_df, pretax, cur_yr), excel_cell(is_df, effective_tax, cur_yr)
        )
        is_df.at[diluted_eps, cur_yr] = '={}/{}'.format(
            excel_cell(is_df, net_income, cur_yr),
            excel_cell(is_df, diluted_share_outstanding, cur_yr)
        )

        if isinstance(is_df.at[ebitda, cur_yr], str):  # two rows with label EBITDA frequently
            is_df.at[ebitda, cur_yr] = '={}+{}'.format(
                excel_cell(is_df, dna_expense, cur_yr), excel_cell(is_df, ebit, cur_yr)
            )
        else:
            is_df.at[ebitda, cur_yr] = [
                np.nan, 
            ]
    empty_unmodified(is_df, yrs_to_predict)

    return is_df, cf_df


def process_bs(is_df, bs_df, cf_df, yrs_to_predict):
    """Manipulates balance sheet."""
    # Short-hands
    last_given_yr = bs_df.columns[-1]

    # Insert 4 empty rows
    bs_df = pd.concat(
        [pd.DataFrame({yr: [np.nan] * 4 for yr in bs_df.columns}, index=[np.nan] * 4), bs_df]
    )

    # Balance sheet labels
    st_receivables = searched_label(bs_df.index, "short term receivables")
    cash_st_investments = searched_label(bs_df.index, "cash short term investments")
    inventories = searched_label(bs_df.index, "inventories")
    other_cur_assets = searched_label(bs_df.index, "other current asset")
    total_cur_assets = searched_label(bs_df.index, "total current asset")
    net_property_plant_equipment = searched_label(bs_df.index, "net property plant equipment")
    total_assets = searched_label(bs_df.index, "total assets")
    accounts_payable = searched_label(bs_df.index, "accounts payable")
    other_cur_liabilities = searched_label(bs_df.index, "other current liabilities")
    total_cur_liabilities = searched_label(bs_df.index, "total current liabilities")
    total_liabilities = searched_label(bs_df.index, "total liabilities")
    total_equity = searched_label(bs_df.index, "total equity")
    total_liabilities_n_shareholders_equity = searched_label(
        bs_df.index, "total liabilities shareholders equity"
    )

    # Income statement labels
    sales = searched_label(is_df.index, "total sales")
    cogs = searched_label(is_df.index, "cost of goods sold cogs excl")
    net_income = searched_label(is_df.index, "net income")

    # Cash flow statement labels
    deprec_deplet_n_amort = searched_label(cf_df.index,
                                           "depreciation depletion amortization expense")
    capital_expenditures = searched_label(cf_df.index, "capital expenditures")
    cash_div_paid = searched_label(cf_df.index, "cash dividends paid")
    change_in_capital_stock = searched_label(cf_df.index, "change in capital stock")

    # Add driver rows to balance sheet
    add_empty_row(bs_df)
    bs_df.loc["Driver Ratios"] = np.nan
    # DSO
    bs_df.loc["DSO"] = [
        "={}/'{}'!{}*365".format(
            excel_cell(bs_df, st_receivables, yr), IS, excel_cell(is_df, sales, yr)
        ) for yr in bs_df.columns
    ]
    dso = "DSO"
    # Other current assets growth %
    add_growth_rate_row(bs_df, other_cur_assets, "Other Current Assets Growth %")
    other_cur_assets_growth = "Other Current Assets Growth %"
    # DPO
    add_empty_row(bs_df)
    bs_df.loc["DPO"] = [
        "={}/'{}'!{}*366".format(
            excel_cell(bs_df, accounts_payable, yr), IS, excel_cell(is_df, cogs, yr)
        ) for yr in bs_df.columns
    ]
    dpo = "DPO"
    # Miscellaneous Current Liabilities Growth %
    add_growth_rate_row(bs_df, other_cur_liabilities, "Miscellaneous Current Liabilities Growth %")
    misc_cur_liabilities_growth = "Miscellaneous Current Liabilities Growth %"

    # Inventory turnober ratio
    if inventories:
        add_empty_row(bs_df)
        bs_df.loc["Inventory Turnover Ratio"] = np.nan
        bs_df.loc["Inventory Turnover Ratio"].iloc[1:] = [
            "='{}'!{}/({}+{})*2".format(
                IS, excel_cell(is_df, cogs, bs_df.columns[i + 1]),
                excel_cell(bs_df, inventories, bs_df.columns[i]),
                excel_cell(bs_df, inventories, bs_df.columns[i+1])
            ) for i in range(len(bs_df.columns) - 1)
        ]
    inventory_ratio = "Inventory Turnover Ratio"

    # Add driver rows to cash flow statement
    add_empty_row(cf_df)
    cf_df.loc["Driver Ratios"] = np.nan
    # Capital Expenditure Revenue Ratio
    cf_df.loc["Capital Expenditure Revenue Ratio"] = [
        "=-{}/'{}'!{}".format(
            excel_cell(cf_df, capital_expenditures, yr), IS, excel_cell(is_df, sales, yr)
        ) for yr in cf_df.columns
    ]
    # Other Funds Net Operating CF Ratio
    net_operating_cf = searched_label(cf_df.index, "net operat cash flow cf")
    initialize_ratio_row(cf_df, searched_label(cf_df.index, "other funds"), net_operating_cf,
                         "Other Funds Net Operating CF Ratio", net_operating_cf)

    # Add prediction years
    for i in range(yrs_to_predict):
        add_yr_column(bs_df)
    for i in range(yrs_to_predict):
        add_yr_column(cf_df)

    # Insert cash balance
    cash_balance_df = pd.DataFrame({yr: np.nan for yr in cf_df.columns}, index=["Cash Balance"])
    cf_df = insert_after(cf_df, cash_balance_df, "net change in tax")
    cash_balance = searched_label(cf_df.index, "cash balance")

    # Inesrt working capital row
    wk_df = pd.DataFrame(
        {
            yr: ['={}-{}'.format(
                excel_cell(bs_df, total_cur_assets, yr),
                excel_cell(bs_df, total_cur_liabilities, yr)
            ), np.nan] for yr in bs_df.columns
        }, index=["Working Capital", np.nan]
    )
    bs_df = insert_before(bs_df, wk_df, "driver ratios")

    # Inesrt balance row
    balance_df = pd.DataFrame(
        {
            yr: ['={}-{}'.format(
                excel_cell(bs_df, total_assets, yr),
                excel_cell(bs_df, total_liabilities_n_shareholders_equity, yr)
            ), np.nan] for yr in bs_df.columns
        }, index=["Balance", np.nan]
    )
    bs_df = insert_before(bs_df, balance_df, "working capital")

    # Calculate driver ratios
    bs_df.loc[dso].iloc[-yrs_to_predict:] = '=' + excel_cell(
        bs_df, dso, bs_df.columns[-yrs_to_predict - 2]
    )
    driver_extend(bs_df, dpo, "avg", last_given_yr, yrs_to_predict)
    driver_extend(bs_df, other_cur_assets_growth, "avg", last_given_yr, yrs_to_predict)
    driver_extend(bs_df, misc_cur_liabilities_growth, "avg", last_given_yr, yrs_to_predict)
    driver_extend(bs_df, inventory_ratio, "avg", last_given_yr, yrs_to_predict)

    # Calculate total liabilities & shareholders' equity

    bs_df.loc[total_liabilities_n_shareholders_equity] = [
        '={}+{}'.format(
            excel_cell(bs_df, total_liabilities, yr),
            excel_cell(bs_df, total_equity, yr)
        ) for yr in bs_df.columns
    ]

    for i in range(yrs_to_predict):
        cur_yr = bs_df.columns[-yrs_to_predict + i]
        prev_yr = bs_df.columns[-yrs_to_predict + i - 1]

        # Calculate variables
        bs_df.at[cash_st_investments, cur_yr] = "='{}'!{}".format(
            CF, excel_cell(cf_df, cash_balance, cur_yr)
        )
        bs_df.at[st_receivables, cur_yr] = "={}/365*'{}'!{}".format(
            excel_cell(bs_df, dso, cur_yr), IS, excel_cell(is_df, sales, cur_yr)
        )
        bs_df.at[other_cur_assets, cur_yr] = '={}*(1+{})'.format(
            excel_cell(bs_df, other_cur_assets, prev_yr),
            excel_cell(bs_df, other_cur_assets_growth, cur_yr)
        )
        bs_df.at[net_property_plant_equipment, cur_yr] = "={}-'{}'!{}-'{}'!{}".format(
            excel_cell(bs_df, net_property_plant_equipment, prev_yr), CF,
            excel_cell(cf_df, deprec_deplet_n_amort, cur_yr), CF,
            excel_cell(cf_df, capital_expenditures, cur_yr)
        )
        bs_df.at[accounts_payable, cur_yr] = "={}/365*'{}'!{}".format(
            excel_cell(bs_df, dpo, cur_yr), IS, excel_cell(is_df, cogs, cur_yr)
        )
        bs_df.at[other_cur_liabilities, cur_yr] = "={}*(1+{})".format(
            excel_cell(bs_df, other_cur_liabilities, prev_yr),
            excel_cell(bs_df, misc_cur_liabilities_growth, cur_yr)
        )
        bs_df.at[total_equity, cur_yr] = "={}+'{}'!{}+'{}'!{}+'{}'!{}".format(
            excel_cell(bs_df, total_equity, prev_yr), CF,
            excel_cell(cf_df, change_in_capital_stock, cur_yr), IS,
            excel_cell(is_df, net_income, cur_yr), CF, excel_cell(cf_df, cash_div_paid, cur_yr)
        )

        # Sums
        bs_df.at[total_cur_assets, cur_yr] = '=' + sum_formula(
            bs_df, total_cur_assets, cur_yr
        )
        bs_df.at[total_assets, cur_yr] = '={}+{}'.format(
            excel_cell(bs_df, total_cur_assets, cur_yr),
            sum_formula(bs_df, total_assets, cur_yr)
        )
        bs_df.at[total_cur_liabilities, cur_yr] = '=' + sum_formula(
            bs_df, total_cur_liabilities, cur_yr
        )
        bs_df.at[total_liabilities, cur_yr] = '={}+{}'.format(
            excel_cell(bs_df, total_cur_liabilities, cur_yr),
            sum_formula(bs_df, total_liabilities, cur_yr)
        )

    for label in bs_df.index[:bs_df.index.get_loc(total_liabilities)]:
        if pd.notna(label) and pd.notna(bs_df.loc[label].iloc[0]) \
                           and bs_df.loc[label].iloc[-1] == '0':
            fixed_extend(bs_df, label, 'prev', yrs_to_predict)

    empty_unmodified(bs_df, yrs_to_predict)

    return bs_df, cf_df


def process_cf(is_df, bs_df, cf_df, yrs_to_predict):
    """Manipulates cash flow statement."""
    # Short-hands
    last_given_yr = cf_df.columns[-yrs_to_predict-1]

    # Cash flow statement labels
    net_income_cf = searched_label(cf_df.index, "net income starting line")
    deprec_deplet_n_amort = searched_label(cf_df.index, "depreciation depletion amortization")
    deferred_taxes = searched_label(cf_df.index, "deferred taxes & investment tax credit")
    other_funds = searched_label(cf_df.index, "other funds")
    funds_from_operations = searched_label(cf_df.index, "funds from operations")
    changes_in_working_capital = searched_label(cf_df.index, "changes in working capital")
    net_operating_cf = searched_label(cf_df.index, "net operating cash flow")
    capital_expenditures = searched_label(cf_df.index, "capital expenditures")
    net_asset_acquisition = searched_label(cf_df.index, "net assets from acquisiton")
    fixed_assets_n_businesses_sale = searched_label(cf_df.index,
                                                    "fixed assets & of sale businesses")
    purchase_sale_of_investments = searched_label(cf_df.index, "purchasesale of investments")
    net_investing_cf = searched_label(cf_df.index, "net investing cash flow")
    cash_div_paid = searched_label(cf_df.index, "cash dividends paid")
    change_in_capital_stock = searched_label(cf_df.index, "change in capital stock")
    net_inssuance_reduction_of_debt = searched_label(cf_df.index, "net issuance reduction of debt")
    net_financing_cf = searched_label(cf_df.index, "net financing cash flow")
    net_change_in_cash = searched_label(cf_df.index, "net change in cash")
    capex_ratio = "Capital Expenditure Revenue Ratio"
    other_funds_net_operating_ratio = "Other Funds Net Operating CF Ratio"

    # Income statement labels
    sales = searched_label(is_df.index, "total sales")
    deprec_amort_expense = searched_label(is_df.index, "depreciation amortization expense")
    net_income_is = searched_label(is_df.index, "net income")
    diluted_share_outstanding = searched_label(is_df.index, "diluted shares outstanding")
    div_per_share = searched_label(is_df.index, "dividends per share")

    # Balance sheet labels
    other_cur_assets = searched_label(bs_df.index, "other current assets")
    other_cur_liabilities = searched_label(bs_df.index, "other current liabilities")
    cash_st_investments = searched_label(bs_df.index, "cash short term investments")
    st_receivables = searched_label(bs_df.index, "short term receivables")
    accounts_payable = searched_label(bs_df.index, "accounts payable")
    lt_debt = searched_label(bs_df.index, "long term debt")

    # Insert cash balance
    cf_df.loc["Cash Balance"].iloc[-yrs_to_predict - 1:] = [
        "='{}'!{}".format(BS, excel_cell(bs_df, cash_st_investments, last_given_yr))
    ] + [
        '={}+{}'.format(
            excel_cell(cf_df, "Cash Balance", cf_df.columns[-yrs_to_predict + i - 1]),
            excel_cell(cf_df, net_change_in_cash, cf_df.columns[-yrs_to_predict + i])
        ) for i in range(yrs_to_predict)
    ]

    # Add levered free CF row
    cf_df.loc["Levered Free Cash Flow"] = [
        '={}+{}'.format(
            excel_cell(cf_df, net_operating_cf, yr),
            excel_cell(cf_df, capital_expenditures, yr)
        ) for yr in cf_df.columns
    ]
    levered_free_cf = "Levered Free Cash Flow"

    # Add levered free CF row growth %
    add_growth_rate_row(cf_df, levered_free_cf, "Levered Free Cash Flow Growth %")
    levered_free_cf_growth = "Levered Free Cash Flow Growth %"

    # Calculate driver ratios
    driver_extend(cf_df, capex_ratio, "avg", last_given_yr, yrs_to_predict)
    driver_extend(cf_df, other_funds_net_operating_ratio, "avg", last_given_yr, yrs_to_predict)

    # Calculate fixed variables
    fixed_extend(cf_df, deferred_taxes, "zero", yrs_to_predict)
    fixed_extend(cf_df, other_funds, "zero", yrs_to_predict)
    fixed_extend(cf_df, net_asset_acquisition, "zero", yrs_to_predict)
    fixed_extend(cf_df, fixed_assets_n_businesses_sale, "zero", yrs_to_predict)
    fixed_extend(cf_df, purchase_sale_of_investments, "zero", yrs_to_predict)
    fixed_extend(cf_df, change_in_capital_stock, "prev", yrs_to_predict)

    # Calculate net operating CF
    cf_df.loc[net_operating_cf] = [
        '={}+{}'.format(
            excel_cell(cf_df, funds_from_operations, yr),
            excel_cell(cf_df, changes_in_working_capital, yr)
        ) for yr in cf_df.columns
    ]

    # Calculate net investing CF
    cf_df.loc[net_investing_cf] = [
        '=' + sum_formula(cf_df, net_investing_cf, yr) for yr in cf_df.columns
    ]

    # Calcualate net financing CF
    cf_df.loc[net_financing_cf] = [
        '=' + sum_formula(cf_df, net_financing_cf, yr) for yr in cf_df.columns
    ]

    # Calculate net change in cash
    cf_df.loc[net_change_in_cash] = [
        '={}+{}+{}'.format(
            excel_cell(cf_df, net_operating_cf, yr), excel_cell(cf_df, net_investing_cf, yr),
            excel_cell(cf_df, net_financing_cf, yr)
        ) for yr in cf_df.columns
    ]

    for i in range(yrs_to_predict):
        cur_yr = is_df.columns[-yrs_to_predict + i]
        prev_yr = is_df.columns[-yrs_to_predict + i - 1]

        # Calculate variables
        cf_df.at[net_income_cf, cur_yr] = "='{}'!{}".format(
            IS, excel_cell(is_df, net_income_is, cur_yr)
        )
        cf_df.at[deprec_deplet_n_amort, cur_yr] = "='{}'!{}".format(
            IS, excel_cell(is_df, deprec_amort_expense, cur_yr)
        )

        cf_df.at[funds_from_operations, cur_yr] = '=' + sum_formula(
            cf_df, funds_from_operations, cur_yr
        )

        cf_df.at[changes_in_working_capital, cur_yr] = "=SUM('{}'!{}:{})-SUM('{}'!{}:{})".format(
            BS, excel_cell(bs_df, st_receivables, prev_yr),
            excel_cell(bs_df, other_cur_assets, prev_yr),
            BS, excel_cell(bs_df, st_receivables, cur_yr),
            excel_cell(bs_df, other_cur_assets, cur_yr)
        )
        cf_df.at[changes_in_working_capital, cur_yr] += "+SUM('{}'!{}:{})-SUM('{}'!{}:{})".format(
            BS, excel_cell(bs_df, accounts_payable, cur_yr),
            excel_cell(bs_df, other_cur_liabilities, cur_yr),
            BS, excel_cell(bs_df, accounts_payable, prev_yr),
            excel_cell(bs_df, other_cur_liabilities, prev_yr)
        )

        cf_df.at[capital_expenditures, cur_yr] = "=-'{}'!{}*{}".format(
            IS, excel_cell(is_df, sales, cur_yr), excel_cell(cf_df, capex_ratio, cur_yr)
        )
        cf_df.at[cash_div_paid, cur_yr] = "=-'{}'!{}*'{}'!{}".format(
            IS, excel_cell(is_df, diluted_share_outstanding, cur_yr),
            IS, excel_cell(is_df, div_per_share, cur_yr)
        )
        cf_df.at[net_inssuance_reduction_of_debt, cur_yr] = "='{}'!{}-'{}'!{}".format(
            BS, excel_cell(bs_df, lt_debt, cur_yr),
            BS, excel_cell(bs_df, lt_debt, prev_yr)
        )
    empty_unmodified(cf_df, yrs_to_predict)

    return cf_df


def main():
    """Main."""
    income_statement = pd.read_excel("asset/{} IS.xlsx".format(NAME), header=4,
                                     index_col=0)
    balance_sheet = pd.read_excel("asset/{} BS.xlsx".format(NAME), header=4, index_col=0)
    cash_flow = pd.read_excel("asset/{} CF.xlsx".format(NAME), header=4, index_col=0)
    market_cap = pd.read_excel("asset/{} MKT.xlsx".format(NAME), index_col=0)

    income_statement, _ = preprocess(income_statement)
    balance_sheet, _ = preprocess(balance_sheet)
    cash_flow, fye = preprocess(cash_flow)

    is_unit, bs_unit = get_unit(income_statement), get_unit(balance_sheet)
    cf_unit, mkt_unit = get_unit(cash_flow), get_unit(market_cap)

    # FIXME
    if is_unit != bs_unit and bs_unit != cf_unit:
        print("Different units.")
        exit(1)
    if mkt_unit != is_unit:
        if mkt_unit == 'm':
            multiply_market = 0.001
        else:  # FIXME
            print("Market unit bigger")
            exit(1)
    else:
        multiply_market = 1

    # Slices of data
    is_search = income_statement.index.get_loc(searched_label(income_statement.index, "EBITDA"))
    if isinstance(is_search, int):
        income_statement = income_statement[:is_search + 1]
    else:
        income_statement = income_statement[:np.where(is_search)[-1][-1] + 1]
    bs_search = balance_sheet.index.get_loc(searched_label(
        balance_sheet.index, "total liabilities & shareholder equity")
    )
    if isinstance(bs_search, int):
        balance_sheet = balance_sheet[:bs_search + 1]
    else:
        balance_sheet = balance_sheet[:np.where(bs_search)[-1][-1] + 1]
    cf_search = cash_flow.index.get_loc(searched_label(cash_flow.index, "net change in cash"))
    if isinstance(cf_search, int):
        cash_flow = cash_flow[:cf_search + 1]
    else:
        cash_flow = cash_flow[:np.where(cf_search)[-1][-1] + 1]

    # FIXME temporary parameters
    growth_rates = [0.5, 0.5, 0.5, 0.5, 0.5]
    yrs_to_predict = len(growth_rates)

    # FIXME diluted shares outstanding
    income_statement.loc["Diluted Shares Outstanding"] = np.nan
    income_statement.loc["Diluted Shares Outstanding"][-1] = market_cap.loc[
        searched_label(market_cap.index, "fully diluted equity capitalization")
    ][0] * multiply_market

    # Cast year data type
    income_statement.columns = income_statement.columns.astype(int)
    balance_sheet.columns = balance_sheet.columns.astype(int)
    cash_flow.columns = cash_flow.columns.astype(int)

    income_statement, cash_flow = process_is(income_statement, cash_flow, growth_rates,
                                             yrs_to_predict)
    balance_sheet, cash_flow = process_bs(income_statement, balance_sheet, cash_flow,
                                          yrs_to_predict)
    cash_flow = process_cf(income_statement, balance_sheet, cash_flow, yrs_to_predict)

    # Stylize excel sheets and output excel
    wb = openpyxl.Workbook()
    wb['Sheet'].title = IS
    wb.create_sheet(BS)
    wb.create_sheet(CF)
    for r in dataframe_to_rows(income_statement):
        wb[IS].append(r)
    for r in dataframe_to_rows(balance_sheet):
        wb[BS].append(r)
    for r in dataframe_to_rows(cash_flow):
        wb[CF].append(r)
    style_ws(wb[IS], IS, income_statement, balance_sheet, cash_flow, fye, is_unit)
    style_ws(wb[BS], BS, income_statement, balance_sheet, cash_flow, fye, is_unit)
    style_ws(wb[CF], CF, income_statement, balance_sheet, cash_flow, fye, is_unit)
    wb.save("output.xlsx")


if __name__ == "__main__":
    main()
