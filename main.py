"""Main program's implementation."""
import os
import openpyxl
import numpy as np
import pandas as pd
from shutil import rmtree
from datetime import datetime
from openpyxl.utils.dataframe import dataframe_to_rows
from style import style_ws
from helper import *

NAME = [
    "NFLX", "AAPL", "PG", "ADS", "AMGN", "AMZN", "CBRE", "COST", "CVX", "DAL", "FB",
    "GOOGL", "MMM", "NKE", "QCOM", "T", "TRIP"
] # GS
IS = "Income Statement"
BS = "Balance Sheet"
CF = "Cashflow Statement"
YRS_TO_CONSIDER = 5

"""
TODO:
    Customize number of years to consider.
    Optimize searched label.
"""

class ValuationMachine:
    def __init__(self, name, growth_rates):
        self.name = name
        self.growth_rates = growth_rates
        self.yrs_to_predict = len(growth_rates)

    def read(self):
        self.is_df = pd.read_excel("asset/{} IS.xlsx".format(self.name), header=4, index_col=0)
        self.bs_df = pd.read_excel("asset/{} BS.xlsx".format(self.name), header=4, index_col=0)
        self.cf_df = pd.read_excel("asset/{} CF.xlsx".format(self.name), header=4, index_col=0)
        self.mkt_df = pd.read_excel("asset/{} MKT.xlsx".format(self.name), index_col=0)

    def preprocess(self):
        self.fye = self.is_df.columns[1]
        self.is_df = preprocess(self.is_df, YRS_TO_CONSIDER)
        self.bs_df = preprocess(self.bs_df, YRS_TO_CONSIDER)
        self.cf_df = preprocess(self.cf_df, YRS_TO_CONSIDER)

    def get_units(self):
        self.is_unit, self.bs_unit = get_unit(self.is_df), get_unit(self.bs_df)
        self.cf_unit, self.mkt_unit = get_unit(self.cf_df), get_unit(self.mkt_df)

        # FIXME
        if self.is_unit != self.bs_unit or self.bs_unit != self.cf_unit:
            print("{} is:{}, bs:{}, cf:{}".format(self.name, self.is_unit, self.bs_unit, self.cf_unit))
    
        if self.mkt_unit != self.is_unit:
            if self.mkt_unit == 'm':
                self.mkt_multiplier = 0.001
            else:  # FIXME
                print("Market unit bigger")
                exit(1)
        else:
            self.mkt_multiplier = 1

        if self.is_unit != self.bs_unit:  # assume CF and IS have the same units
            if self.bs_unit == 'b':
                self.extra_bs = "/1000"
                self.extra_cf = "*1000"
            else:
                self.extra_bs = "*1000"
                self.extra_cf = "/1000"
        else:
            self.extra_bs = ""
            self.extra_cf = ""


    def slice_data(self):
        if searched_label(self.is_df.index, "EBITDA"):
            is_search = self.is_df.index.get_loc(searched_label(self.is_df.index, "EBITDA"))
        else:
            is_search = self.is_df.index.get_loc(
                searched_label(self.is_df.index, "dividends per share")
            )
        if isinstance(is_search, int):
            self.is_df = self.is_df[:is_search + 1]
        else:
            self.is_df = self.is_df[:np.where(is_search)[-1][-1] + 1]
        bs_search = self.bs_df.index.get_loc(searched_label(
            self.bs_df.index, "total liabilities & shareholder equity")
        )
        if isinstance(bs_search, int):
            self.bs_df = self.bs_df[:bs_search + 1]
        else:
            self.bs_df = self.bs_df[:np.where(bs_search)[-1][-1] + 1]
        cf_search = self.cf_df.index.get_loc(
            searched_label(self.cf_df.index, "net change in cash")
        )
        if isinstance(cf_search, int):
            self.cf_df = self.cf_df[:cf_search + 1]
        else:
            self.cf_df = self.cf_df[:np.where(cf_search)[-1][-1] + 1]

    def style(self):
        # Stylize excel sheets and output excel
        wb = openpyxl.Workbook()
        wb['Sheet'].title = IS
        wb.create_sheet(BS)
        wb.create_sheet(CF)
        for r in dataframe_to_rows(self.is_df):
            wb[IS].append(r)
        for r in dataframe_to_rows(self.bs_df):
            wb[BS].append(r)
        for r in dataframe_to_rows(self.cf_df):
            wb[CF].append(r)
        style_ws(wb[IS], IS, self.is_df, self.bs_df, self.cf_df, self.fye, self.is_unit)
        style_ws(wb[BS], BS, self.is_df, self.bs_df, self.cf_df, self.fye, self.bs_unit)
        style_ws(wb[CF], CF, self.is_df, self.bs_df, self.cf_df, self.fye, self.cf_unit)
        wb.save("output/output_{}_{}.xlsx".format(self.name, datetime.now().strftime('%H-%M-%S')))

    def process_is(self):
        """Manipulates income statement."""
        # Short-hands
        yrs_to_predict = self.yrs_to_predict
        last_given_yr = self.is_df.columns[-1]

        # Diluted shares outstanding
        self.is_df.loc["Diluted Shares Outstanding"] = np.nan
        self.is_df.loc["Diluted Shares Outstanding"].iloc[-1] = self.mkt_df.loc[
            searched_label(self.mkt_df.index, "fully diluted equity capitalization")
        ][0] * self.mkt_multiplier

        # Declare income statement labels
        sales = searched_label(self.is_df.index, "total sales")
        cogs = searched_label(self.is_df.index, "cost of goods sold COGS including")
        self.is_df.index = [  # change label name from incl. to excl.
            i if i != cogs else "Cost of Goods Sold (COGS) excl. D&A" \
            for i in list(self.is_df.index)
        ]
        cogs = "Cost of Goods Sold (COGS) excl. D&A"
        gross_income = searched_label(self.is_df.index, "gross income")
        sgna = searched_label(self.is_df.index, "sg&a expense")
        other_expense = searched_label(self.is_df.index, "other expense")
        ebit = searched_label(self.is_df.index, "ebit operating income")
        nonoperating_income = searched_label(self.is_df.index, "nonoperating income net")
        interest_expense = searched_label(self.is_df.index, "interest expense")
        unusual = searched_label(self.is_df.index, "unusual expense")
        income_tax = searched_label(self.is_df.index, "income taxes")
        diluted_eps = searched_label(self.is_df.index, "eps diluted")
        net_income = searched_label(self.is_df.index, "net income")
        div_per_share = searched_label(self.is_df.index, "dividends per share")
        ebitda = searched_label(self.is_df.index, "ebitda")
        diluted_share_outstanding = "Diluted Shares Outstanding"

        # Drop EBITDA if exists
        if ebitda:
            self.is_df.drop(ebitda, inplace=True)

        # Insert pretax income row before income taxes
        pretax_df = pd.DataFrame(
            {
                yr: '={}+{}-{}-{}'.format(
                    excel_cell(self.is_df, ebit, yr),
                    excel_cell(self.is_df, nonoperating_income, yr),
                    excel_cell(self.is_df, interest_expense, yr),
                    excel_cell(self.is_df, unusual, yr)
                ) for yr in self.is_df.columns
            }, index=["Pretax Income"]
        )
        self.is_df = insert_before(self.is_df, pretax_df, income_tax)
        pretax = "Pretax Income"

        # Insert depreciation & amortization expense before SG&A expense
        dna_expense_df = pd.DataFrame(
            {
                yr: "='{}'!".format(CF) + excel_cell(
                    self.cf_df, searched_label(
                        self.cf_df.index, "depreciation depletion & amortization expense"
                    ), yr
                ) for yr in self.is_df.columns
            }, index=["Depreciation & Amortization Expense"]
        )
        self.is_df = insert_before(self.is_df, dna_expense_df, sgna)
        dna_expense = "Depreciation & Amortization Expense"

        # Write formulas for COGS excl. D&A
        self.is_df.loc[cogs] = [
            '={}-{}'.format(self.is_df.at[cogs, yr], excel_cell(self.is_df, dna_expense, yr))
            for yr in self.is_df.columns
        ]

        # Add driver rows to income statement
        add_empty_row(self.is_df)
        self.is_df.loc["Driver Ratios"] = np.nan
        add_growth_rate_row(self.is_df, sales, "Sales Growth %")
        sales_growth = "Sales Growth %"
        add_empty_row(self.is_df)
        initialize_ratio_row(self.is_df, cogs, sales, "COGS Sales Ratio")
        cogs_ratio = "COGS Sales Ratio"
        add_empty_row(self.is_df)
        initialize_ratio_row(self.is_df, sgna, sales, "SG&A Sales Ratio")
        sgna_ratio = "SG&A Sales Ratio"
        add_empty_row(self.is_df)
        initialize_ratio_row(self.is_df, dna_expense, sales, "D&A Sales Ratio")
        dna_ratio = "D&A Sales Ratio"
        add_empty_row(self.is_df)
        initialize_ratio_row(self.is_df, unusual, ebit, "Unusual Expense EBIT Ratio")
        unusual_ratio = "Unusual Expense EBIT Ratio"
        add_empty_row(self.is_df)
        initialize_ratio_row(self.is_df, income_tax, pretax, "Effective Tax Rate")
        effective_tax = "Effective Tax Rate"

        # Add prediction years
        for i in range(yrs_to_predict):
            add_yr_column(self.is_df)

        # Append growth rates to driver row
        self.is_df.loc[sales_growth].iloc[-yrs_to_predict:] = self.growth_rates

        # Write formulas for EBITDA
        ebitda_df = pd.DataFrame(
            {
                yr: '={}+{}'.format(
                    excel_cell(self.is_df, dna_expense, yr), excel_cell(self.is_df, ebit, yr)
                ) for yr in self.is_df.columns
            }, index=[ebitda]
        )
        self.is_df = insert_after(self.is_df, ebitda_df, div_per_share)

        # Write formulas for driver ratios
        initialize_ratio_row(self.is_df, div_per_share, diluted_eps, "Dividend Payout Ratio")
        initialize_ratio_row(self.is_df, ebitda, sales, "EBITDA Margin", sales_growth)
        self.is_df.loc[dna_ratio].iloc[-yrs_to_predict:] = self.is_df.loc[dna_ratio, last_given_yr]
        driver_extend(self.is_df, cogs_ratio, "round", last_given_yr, yrs_to_predict)
        driver_extend(self.is_df, sgna_ratio, "round", last_given_yr, yrs_to_predict)
        driver_extend(self.is_df, unusual_ratio, "avg", last_given_yr, yrs_to_predict)
        driver_extend(self.is_df, diluted_share_outstanding, "avg",
                      last_given_yr, yrs_to_predict, 3)
        driver_extend(self.is_df, effective_tax, "avg", last_given_yr, yrs_to_predict)

        # Calculate fixed variables
        fixed_extend(self.is_df, other_expense, "prev", yrs_to_predict)
        fixed_extend(self.is_df, nonoperating_income, "prev", yrs_to_predict)
        fixed_extend(self.is_df, interest_expense, "prev", yrs_to_predict)
        fixed_extend(self.is_df, div_per_share, "prev", yrs_to_predict)  # FIXME
        fixed_extend(self.is_df, diluted_share_outstanding, "prev", yrs_to_predict)

        # Write formula for net income
        self.is_df.loc[net_income] = [
            '={}-{}'.format(
                excel_cell(self.is_df, pretax, yr),
                excel_cell(self.is_df, income_tax, yr)
            ) for yr in self.is_df.columns
        ]

        for i in range(yrs_to_predict):
            cur_yr = self.is_df.columns[-yrs_to_predict + i]
            prev_yr = self.is_df.columns[-yrs_to_predict + i - 1]

            # Write formulas
            self.is_df.at[sales, cur_yr] = '={}*(1+{})'.format(
                excel_cell(self.is_df, sales, prev_yr),
                excel_cell(self.is_df, sales_growth, cur_yr)
            )
            self.is_df.at[cogs, cur_yr] = '={}*{}'.format(
                excel_cell(self.is_df, sales, cur_yr), excel_cell(self.is_df, cogs_ratio, cur_yr)
            )
            self.is_df.at[gross_income, cur_yr] = '={}-{}'.format(
                excel_cell(self.is_df, sales, cur_yr), excel_cell(self.is_df, cogs, cur_yr)
            )
            self.is_df.at[dna_expense, cur_yr] = '={}*{}'.format(
                excel_cell(self.is_df, sales, cur_yr), excel_cell(self.is_df, dna_ratio, cur_yr)
            )
            self.is_df.at[sgna, cur_yr] = '={}*{}'.format(
                excel_cell(self.is_df, sales, cur_yr), excel_cell(self.is_df, sgna_ratio, cur_yr)
            )
            self.is_df.at[ebit, cur_yr] = '={}-{}'.format(
                excel_cell(self.is_df, gross_income, cur_yr),
                sum_formula(self.is_df, ebit, cur_yr, gross_income, 1)
            )
            self.is_df.at[unusual, cur_yr] = '={}*{}'.format(
                excel_cell(self.is_df, ebit, cur_yr),
                excel_cell(self.is_df, unusual_ratio, cur_yr)
            )
            self.is_df.at[pretax, cur_yr] = '={}+{}-{}-{}'.format(
                excel_cell(self.is_df, ebit, cur_yr),
                excel_cell(self.is_df, nonoperating_income, cur_yr),
                excel_cell(self.is_df, interest_expense, cur_yr),
                excel_cell(self.is_df, unusual, cur_yr)
            )
            self.is_df.at[income_tax, cur_yr] = '={}*{}'.format(
                 excel_cell(self.is_df, pretax, cur_yr),
                 excel_cell(self.is_df, effective_tax, cur_yr)
            )
            self.is_df.at[diluted_eps, cur_yr] = '={}/{}'.format(
                excel_cell(self.is_df, net_income, cur_yr),
                excel_cell(self.is_df, diluted_share_outstanding, cur_yr)
            )

            if isinstance(self.is_df.at[ebitda, cur_yr], str):  # two rows with label EBITDA
                self.is_df.at[ebitda, cur_yr] = '={}+{}'.format(
                    excel_cell(self.is_df, dna_expense, cur_yr),
                    excel_cell(self.is_df, ebit, cur_yr)
                )
            else:
                self.is_df.at[ebitda, cur_yr] = [
                    np.nan, 
                ]
        empty_unmodified(self.is_df, yrs_to_predict)


    def process_bs(self):
        """Manipulates balance sheet."""
        # Short-hands
        yrs_to_predict = self.yrs_to_predict
        last_given_yr = self.bs_df.columns[-1]

        # Balance sheet labels
        st_receivables = searched_label(self.bs_df.index, "short term receivables")
        cash_st_investments = searched_label(self.bs_df.index, "cash short term investments")
        inventories = searched_label(self.bs_df.index, "inventories")
        other_cur_assets = searched_label(self.bs_df.index, "other current asset")
        total_cur_assets = searched_label(self.bs_df.index, "total current asset")
        net_property_plant_equipment = searched_label(self.bs_df.index,
                                                      "net property plant equipment")
        total_assets = searched_label(self.bs_df.index, "total assets")
        accounts_payable = searched_label(self.bs_df.index, "accounts payable")
        other_cur_liabilities = searched_label(self.bs_df.index, "other current liabilities")
        total_cur_liabilities = searched_label(self.bs_df.index, "total current liabilities")
        total_liabilities = searched_label(self.bs_df.index, "total liabilities")
        total_equity = searched_label(self.bs_df.index, "total equity")
        total_liabilities_n_shareholders_equity = searched_label(
            self.bs_df.index, "total liabilities shareholders equity"
        )

        # Income statement labels
        sales = searched_label(self.is_df.index, "total sales")
        cogs = searched_label(self.is_df.index, "cost of goods sold cogs excl")
        net_income = searched_label(self.is_df.index, "net income")

        # Cash flow statement labels
        deprec_deplet_n_amort = searched_label(self.cf_df.index,
                                           "depreciation depletion amortization expense")
        capital_expenditures = searched_label(self.cf_df.index, "capital expenditures")
        cash_div_paid = searched_label(self.cf_df.index, "cash dividends paid")
        change_in_capital_stock = searched_label(self.cf_df.index, "change in capital stock")

        # Add driver rows to balance sheet
        add_empty_row(self.bs_df)
        self.bs_df.loc["Driver Ratios"] = np.nan
        # DSO
        if st_receivables:
            self.bs_df.loc["DSO"] = [
                "={}/'{}'!{}{}*365".format(
                    excel_cell(self.bs_df, st_receivables, yr), IS,
                    excel_cell(self.is_df, sales, yr), self.extra_cf
                ) for yr in self.bs_df.columns
            ]
            dso = "DSO"
        # Other current assets growth %
        add_growth_rate_row(self.bs_df, other_cur_assets, "Other Current Assets Growth %")
        other_cur_assets_growth = "Other Current Assets Growth %"
        # DPO
        add_empty_row(self.bs_df)
        self.bs_df.loc["DPO"] = [
            "={}/'{}'!{}{}*366".format(
                excel_cell(self.bs_df, accounts_payable, yr), IS,
                excel_cell(self.is_df, cogs, yr), self.extra_cf
            ) for yr in self.bs_df.columns
        ]
        dpo = "DPO"
        # Miscellaneous Current Liabilities Growth %
        add_growth_rate_row(self.bs_df, other_cur_liabilities,
                            "Miscellaneous Current Liabilities Growth %")
        misc_cur_liabilities_growth = "Miscellaneous Current Liabilities Growth %"

        # Inventory turnober ratio
        if inventories:
            add_empty_row(self.bs_df)
            self.bs_df.loc["Inventory Turnover Ratio"] = np.nan
            self.bs_df.loc["Inventory Turnover Ratio"].iloc[1:] = [
                "='{}'!{}{}/({}+{})*2".format(
                    IS, excel_cell(self.is_df, cogs, self.bs_df.columns[i + 1]), self.extra_bs,
                    excel_cell(self.bs_df, inventories, self.bs_df.columns[i]),
                    excel_cell(self.bs_df, inventories, self.bs_df.columns[i+1])
                ) for i in range(len(self.bs_df.columns) - 1)
            ]
        inventory_ratio = "Inventory Turnover Ratio"

        # Add driver rows to cash flow statement
        add_empty_row(self.cf_df)
        self.cf_df.loc["Driver Ratios"] = np.nan
        # Capital Expenditure Revenue Ratio
        self.cf_df.loc["Capital Expenditure Revenue Ratio"] = [
            "=-{}/'{}'!{}".format(
                excel_cell(self.cf_df, capital_expenditures, yr), IS,
                excel_cell(self.is_df, sales, yr)
            ) for yr in self.cf_df.columns
        ]
        # Other Funds Net Operating CF Ratio
        net_operating_cf = searched_label(self.cf_df.index, "net operat cash flow cf")
        initialize_ratio_row(
            self.cf_df, searched_label(self.cf_df.index, "other funds"),
            net_operating_cf, "Other Funds Net Operating CF Ratio", net_operating_cf
        )

        # Add prediction years
        for i in range(yrs_to_predict):
            add_yr_column(self.bs_df)
        for i in range(yrs_to_predict):
            add_yr_column(self.cf_df)

        # Insert cash balance
        cash_balance_df = pd.DataFrame({yr: np.nan for yr in self.cf_df.columns},
                                       index=["Cash Balance"])
        self.cf_df = insert_after(self.cf_df, cash_balance_df, "net change in tax")
        cash_balance = searched_label(self.cf_df.index, "cash balance")

        # Inesrt working capital row
        wk_df = pd.DataFrame(
            {
                yr: ['={}-{}'.format(
                    excel_cell(self.bs_df, total_cur_assets, yr),
                    excel_cell(self.bs_df, total_cur_liabilities, yr)
                ), np.nan] for yr in self.bs_df.columns
            }, index=["Working Capital", np.nan]
        )
        self.bs_df = insert_before(self.bs_df, wk_df, "driver ratios")

        # Inesrt balance row
        balance_df = pd.DataFrame(
            {
                yr: ['={}-{}'.format(
                    excel_cell(self.bs_df, total_assets, yr),
                    excel_cell(self.bs_df, total_liabilities_n_shareholders_equity, yr)
                ), np.nan] for yr in self.bs_df.columns
            }, index=["Balance", np.nan]
        )   
        self.bs_df = insert_before(self.bs_df, balance_df, "working capital")

        # Calculate driver ratios
        if st_receivables:
            self.bs_df.loc[dso].iloc[-yrs_to_predict:] = '=' + excel_cell(
                self.bs_df, dso, self.bs_df.columns[-yrs_to_predict - 2]
            )
        driver_extend(self.bs_df, dpo, "avg", last_given_yr, yrs_to_predict)
        driver_extend(self.bs_df, other_cur_assets_growth, "avg", last_given_yr, yrs_to_predict)
        driver_extend(self.bs_df, misc_cur_liabilities_growth, "avg",
                      last_given_yr, yrs_to_predict)
        driver_extend(self.bs_df, inventory_ratio, "avg", last_given_yr, yrs_to_predict)

        # Calculate total liabilities & shareholders' equity

        self.bs_df.loc[total_liabilities_n_shareholders_equity] = [
            '={}+{}'.format(
                excel_cell(self.bs_df, total_liabilities, yr),
                excel_cell(self.bs_df, total_equity, yr)
            ) for yr in self.bs_df.columns
        ]

        for i in range(yrs_to_predict):
            cur_yr = self.bs_df.columns[-yrs_to_predict + i]
            prev_yr = self.bs_df.columns[-yrs_to_predict + i - 1]

            # Calculate variables
            self.bs_df.at[cash_st_investments, cur_yr] = "='{}'!{}{}".format(
                CF, excel_cell(self.cf_df, cash_balance, cur_yr), self.extra_bs
            )
            if st_receivables:
                self.bs_df.at[st_receivables, cur_yr] = "={}/365*'{}'!{}{}".format(
                    excel_cell(self.bs_df, dso, cur_yr), IS, excel_cell(self.is_df, sales, cur_yr),
                    self.extra_bs
                )
            self.bs_df.at[other_cur_assets, cur_yr] = '={}*(1+{})'.format(
                excel_cell(self.bs_df, other_cur_assets, prev_yr),
                excel_cell(self.bs_df, other_cur_assets_growth, cur_yr)
            )
            self.bs_df.at[net_property_plant_equipment, cur_yr] = "={}-'{}'!{}{}-'{}'!{}{}".format(
                excel_cell(self.bs_df, net_property_plant_equipment, prev_yr), CF,
                excel_cell(self.cf_df, deprec_deplet_n_amort, cur_yr), self.extra_bs, CF,
                excel_cell(self.cf_df, capital_expenditures, cur_yr), self.extra_bs
            )
            self.bs_df.at[accounts_payable, cur_yr] = "={}/365*'{}'!{}{}".format(
                excel_cell(self.bs_df, dpo, cur_yr), IS, excel_cell(self.is_df, cogs, cur_yr),
                self.extra_bs
            )
            self.bs_df.at[other_cur_liabilities, cur_yr] = "={}*(1+{})".format(
                excel_cell(self.bs_df, other_cur_liabilities, prev_yr),
                excel_cell(self.bs_df, misc_cur_liabilities_growth, cur_yr)
            )
            self.bs_df.at[total_equity, cur_yr] = "={}+'{}'!{}{}+'{}'!{}{}+'{}'!{}{}".format(
                excel_cell(self.bs_df, total_equity, prev_yr), CF,
                excel_cell(self.cf_df, change_in_capital_stock, cur_yr), self.extra_bs, IS,
                excel_cell(self.is_df, net_income, cur_yr), self.extra_bs, CF,
                excel_cell(self.cf_df, cash_div_paid, cur_yr), self.extra_bs
            )

            # Sums
            self.bs_df.at[total_cur_assets, cur_yr] = '=' + sum_formula(
                self.bs_df, total_cur_assets, cur_yr
            )
            self.bs_df.at[total_assets, cur_yr] = '={}+{}'.format(
                excel_cell(self.bs_df, total_cur_assets, cur_yr),
                sum_formula(self.bs_df, total_assets, cur_yr)
            )
            self.bs_df.at[total_cur_liabilities, cur_yr] = '=' + sum_formula(
                self.bs_df, total_cur_liabilities, cur_yr
            )
            self.bs_df.at[total_liabilities, cur_yr] = '={}+{}'.format(
                excel_cell(self.bs_df, total_cur_liabilities, cur_yr),
                sum_formula(self.bs_df, total_liabilities, cur_yr)
            )

        for label in self.bs_df.index[:self.bs_df.index.get_loc(total_liabilities)]:
            if pd.notna(label) and pd.notna(self.bs_df.loc[label].iloc[0]) \
                               and self.bs_df.loc[label].iloc[-1] == '0':
                fixed_extend(self.bs_df, label, 'prev', yrs_to_predict)

        empty_unmodified(self.bs_df, yrs_to_predict)


    def process_cf(self):
        """Manipulates cash flow statement."""
        # Short-hands
        yrs_to_predict = self.yrs_to_predict
        last_given_yr = self.cf_df.columns[-yrs_to_predict-1]

        # Cash flow statement labels
        net_income_cf = searched_label(self.cf_df.index, "net income starting line")
        deprec_deplet_n_amort = searched_label(self.cf_df.index,
                                               "depreciation depletion amortization")
        deferred_taxes = searched_label(self.cf_df.index,
                                        "deferred taxes & investment tax credit")
        other_funds = searched_label(self.cf_df.index, "other funds")
        funds_from_operations = searched_label(self.cf_df.index, "funds from operations")
        changes_in_working_capital = searched_label(self.cf_df.index,
                                                    "changes in working capital")
        net_operating_cf = searched_label(self.cf_df.index, "net operating cash flow")
        capital_expenditures = searched_label(self.cf_df.index, "capital expenditures")
        net_asset_acquisition = searched_label(self.cf_df.index, "net assets from acquisiton")
        fixed_assets_n_businesses_sale = searched_label(self.cf_df.index,
                                                        "fixed assets & of sale businesses")
        purchase_sale_of_investments = searched_label(self.cf_df.index,
                                                      "purchasesale of investments")
        net_investing_cf = searched_label(self.cf_df.index, "net investing cash flow")
        cash_div_paid = searched_label(self.cf_df.index, "cash dividends paid")
        change_in_capital_stock = searched_label(self.cf_df.index, "change in capital stock")
        net_inssuance_reduction_of_debt = searched_label(self.cf_df.index,
                                                         "net issuance reduction of debt")
        net_financing_cf = searched_label(self.cf_df.index, "net financing cash flow")
        net_change_in_cash = searched_label(self.cf_df.index, "net change in cash")
        capex_ratio = "Capital Expenditure Revenue Ratio"   
        other_funds_net_operating_ratio = "Other Funds Net Operating CF Ratio"

        # Income statement labels
        sales = searched_label(self.is_df.index, "total sales")
        deprec_amort_expense = searched_label(self.is_df.index,
                                              "depreciation amortization expense")
        net_income_is = searched_label(self.is_df.index, "net income")
        diluted_share_outstanding = searched_label(self.is_df.index, "diluted shares outstanding")
        div_per_share = searched_label(self.is_df.index, "dividends per share")

        # Balance sheet labels
        other_cur_assets = searched_label(self.bs_df.index, "other current assets")
        other_cur_liabilities = searched_label(self.bs_df.index, "other current liabilities")
        cash_st_investments = searched_label(self.bs_df.index, "cash short term investments")
        st_receivables = searched_label(self.bs_df.index, "short term receivables")
        accounts_payable = searched_label(self.bs_df.index, "accounts payable")
        lt_debt = searched_label(self.bs_df.index, "long term debt")

        # Insert cash balance
        self.cf_df.loc["Cash Balance"].iloc[-yrs_to_predict - 1:] = [
            "='{}'!{}{}".format(
                BS, excel_cell(self.bs_df, cash_st_investments, last_given_yr), self.extra_cf
            )
        ] + [
            '={}+{}'.format(
                excel_cell(
                    self.cf_df, "Cash Balance", self.cf_df.columns[-yrs_to_predict + i - 1]
                ),
                excel_cell(
                    self.cf_df, net_change_in_cash, self.cf_df.columns[-yrs_to_predict + i]
                )
            ) for i in range(yrs_to_predict)
        ]

        # Add levered free CF row
        self.cf_df.loc["Levered Free Cash Flow"] = [
            '={}+{}'.format(
                excel_cell(self.cf_df, net_operating_cf, yr),
                excel_cell(self.cf_df, capital_expenditures, yr)
            ) for yr in self.cf_df.columns
        ]
        levered_free_cf = "Levered Free Cash Flow"

        # Add levered free CF row growth %
        add_growth_rate_row(self.cf_df, levered_free_cf, "Levered Free Cash Flow Growth %")
        levered_free_cf_growth = "Levered Free Cash Flow Growth %"

        # Calculate driver ratios
        driver_extend(self.cf_df, capex_ratio, "avg", last_given_yr, yrs_to_predict)
        driver_extend(self.cf_df, other_funds_net_operating_ratio, "avg",
                      last_given_yr, yrs_to_predict)

        # Calculate fixed variables
        fixed_extend(self.cf_df, deferred_taxes, "zero", yrs_to_predict)
        fixed_extend(self.cf_df, other_funds, "zero", yrs_to_predict)
        fixed_extend(self.cf_df, net_asset_acquisition, "zero", yrs_to_predict)
        fixed_extend(self.cf_df, fixed_assets_n_businesses_sale, "zero", yrs_to_predict)
        fixed_extend(self.cf_df, purchase_sale_of_investments, "zero", yrs_to_predict)
        fixed_extend(self.cf_df, change_in_capital_stock, "prev", yrs_to_predict)

        # Calculate net operating CF
        self.cf_df.loc[net_operating_cf] = [
            '={}+{}'.format(
                excel_cell(self.cf_df, funds_from_operations, yr),
                excel_cell(self.cf_df, changes_in_working_capital, yr)
            ) for yr in self.cf_df.columns
        ]

        # Calculate net investing CF
        self.cf_df.loc[net_investing_cf] = [
            '=' + sum_formula(self.cf_df, net_investing_cf, yr) for yr in self.cf_df.columns
        ]

        # Calcualate net financing CF
        self.cf_df.loc[net_financing_cf] = [
            '=' + sum_formula(self.cf_df, net_financing_cf, yr) for yr in self.cf_df.columns
        ]

        # Calculate net change in cash
        self.cf_df.loc[net_change_in_cash] = [
            '={}+{}+{}'.format(
                excel_cell(self.cf_df, net_operating_cf, yr),
                excel_cell(self.cf_df, net_investing_cf, yr),
                excel_cell(self.cf_df, net_financing_cf, yr)
            ) for yr in self.cf_df.columns
        ]

        for i in range(yrs_to_predict):
            cur_yr = self.is_df.columns[-yrs_to_predict + i]
            prev_yr = self.is_df.columns[-yrs_to_predict + i - 1]

            # Calculate variables
            self.cf_df.at[net_income_cf, cur_yr] = "='{}'!{}".format(
                IS, excel_cell(self.is_df, net_income_is, cur_yr)
            )
            self.cf_df.at[deprec_deplet_n_amort, cur_yr] = "='{}'!{}".format(
                IS, excel_cell(self.is_df, deprec_amort_expense, cur_yr)
            )

            self.cf_df.at[funds_from_operations, cur_yr] = '=' + sum_formula(
                self.cf_df, funds_from_operations, cur_yr
            )

            self.cf_df.at[changes_in_working_capital, cur_yr] = "=SUM('{}'!{}:{}){}".format(
                BS, excel_cell(
                    self.bs_df, self.bs_df.index[self.bs_df.index.get_loc(cash_st_investments) + 1],
                    prev_yr
                ), excel_cell(self.bs_df, other_cur_assets, prev_yr), self.extra_cf
            )
            self.cf_df.at[changes_in_working_capital, cur_yr] += "-SUM('{}'!{}:{}){}".format(
                BS, excel_cell(
                    self.bs_df, self.bs_df.index[self.bs_df.index.get_loc(cash_st_investments) + 1],
                    cur_yr
                ), excel_cell(self.bs_df, other_cur_assets, cur_yr), self.extra_cf
            )
            self.cf_df.at[changes_in_working_capital, cur_yr] += "+SUM('{}'!{}:{}){}".format(
                BS, excel_cell(self.bs_df, accounts_payable, cur_yr),
                excel_cell(self.bs_df, other_cur_liabilities, cur_yr), self.extra_cf
            )
            self.cf_df.at[changes_in_working_capital, cur_yr] += "-SUM('{}'!{}:{}){}".format(
                BS, excel_cell(self.bs_df, accounts_payable, prev_yr),
                excel_cell(self.bs_df, other_cur_liabilities, prev_yr), self.extra_cf
            )

            self.cf_df.at[capital_expenditures, cur_yr] = "=-'{}'!{}*{}".format(
                IS, excel_cell(self.is_df, sales, cur_yr),
                excel_cell(self.cf_df, capex_ratio, cur_yr)
            )
            self.cf_df.at[cash_div_paid, cur_yr] = "=-'{}'!{}*'{}'!{}".format(
                IS, excel_cell(self.is_df, diluted_share_outstanding, cur_yr),
                IS, excel_cell(self.is_df, div_per_share, cur_yr)
            )   
            self.cf_df.at[net_inssuance_reduction_of_debt, cur_yr] = "='{}'!{}{}-'{}'!{}{}".format(
                BS, excel_cell(self.bs_df, lt_debt, cur_yr), self.extra_cf,
                BS, excel_cell(self.bs_df, lt_debt, prev_yr), self.extra_cf
            )
        empty_unmodified(self.cf_df, yrs_to_predict)


def main():
    """Call all methods."""
    if os.path.exists('output'):
        rmtree('output')
    os.mkdir('output')

    growth_rates = [0.5, 0.5, 0.5, 0.5, 0.5]
    for i in NAME:
        vm = ValuationMachine(i, growth_rates)
        vm.read()
        vm.preprocess()
        vm.get_units()
        vm.slice_data()
        vm.process_is()
        vm.process_bs()
        vm.process_cf()
        vm.style()


if __name__ == "__main__":
    main()
