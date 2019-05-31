import os
import decimal
from decimal import getcontext, Decimal
import numpy as np
import pandas as pd

getcontext().prec = 4

us_10y_treasury_bond = 0.0268 #as of 12/31/2018 from treasury.gov
sp500_index = 2506.85 #as of 12/31/2019 from us.spindices.com
db = 136.65 #dividend and buyback as for base year
dbgr = 0.0412 #compounded annual rate over next 5 years for dividend and buyback growth rate based on analyst estimate
t_m = 0.25 #marginal tax rate
d_e = 0.00 #need to calculate

def cost_of_debt():
	rating = raw_input("Please provide the bond rating for the firm: ")
	default_spread = 0.00
	print rating
	if rating == "D2" or rating == "D":
		default_spread = 0.1938
	elif rating == "C2" or rating == "C":
		default_spread = 0.1454
	elif rating == "Ca2" or rating == "CC":
		default_spread = 0.1108
	elif rating == "Caa" or rating == "CCC":
		default_spread = 0.09
	elif rating == "B3" or rating == "B-":
		default_spread = 0.066
	elif rating == "B2" or rating == "B":
		default_spread = 0.054
	elif rating == "B1" or rating == "B+":
		default_spread = 0.045
	elif rating == "Ba2" or rating == "BB":
		default_spread = 0.036
	elif rating == "Ba1" or rating == "BB+":
		default_spread = 0.03
	elif rating == "Baa2" or rating == "BBB":
		default_spread = 0.02
	elif rating == "A3" or rating == "A-":
		default_spread = 0.0156
	elif rating == "A2" or rating == "A":
		default_spread = 0.0138
	elif rating == "A1" or rating == "A+":
		default_spread = 0.0125
	elif rating == "Aa2" or rating == "AA":
		default_spread = 0.01
	elif rating == "Aaa" or rating == "AAA":
		default_spread = 0.0075
	cod = (us_10y_treasury_bond + default_spread) * (1 - t_m)
	print "cost of debt:", round(cod * 100, 2), "%"
	return cod


def wacc():
    # implied market premium for s&p 500
    mrp = 0.05  # base testing point as of 2019
    while True:
        CF1 = db * (1 + dbgr) / (1 + mrp)
        CF2 = db * (1 + dbgr) ** 2 / (1 + mrp) ** 2
        CF3 = db * (1 + dbgr) ** 3 / (1 + mrp) ** 3
        CF4 = db * (1 + dbgr) ** 4 / (1 + mrp) ** 4
        CF5 = db * (1 + dbgr) ** 5 / (1 + mrp) ** 5
        TV = db * (1 + dbgr) ** 5 * (1 + us_10y_treasury_bond) / ((mrp - us_10y_treasury_bond) * (1 + mrp) ** 5)
        sp500_simulated = CF1 + CF2 + CF3 + CF4 + CF5 + TV
        mrp += 0.0001
        if sp500_simulated < sp500_index:
            mrp -= 0.0001
            break
    us_imp = mrp - us_10y_treasury_bond # US market risk premium

    #mkt value calculation
    debt_sum = pd.read_excel('NVIDIA/NVIDIA Debt Summary.xlsx', header=4, index_col=0, na_filter=False)
    mkt_debt = round(debt_sum['Mkt Val  (MM)'].T.sum() / 2, 4)
    print "mkt_debt:", mkt_debt
    bs = pd.read_excel('NVIDIA/NVIDIA Mkt Cap.xlsx', header=2, index_col=0, usecols = "A:C")
    mkt_equity = round(bs.loc['Common Stock', 'Mkt Cap'], 4)
    print "mkt_equity:", mkt_equity
    d_e = (mkt_debt / (mkt_debt + mkt_equity)) #operating leases include
    print "d/e ratio:", round(d_e * 100, 2), "%"

    #beta calculation
    beta = pd.read_excel('asset/BETAS.xls', header=9, index_col=0, na_filter=False)
    unlevered_beta = beta.loc['Semiconductor Equip', 'Unlevered beta corrected for cash']  # should be extracted from beta.xlsx
    levered_beta = unlevered_beta * (1 + (1 - t_m) * d_e)
    print "beta:", round(levered_beta, 2)

    #capm calculation
    e_return = us_10y_treasury_bond + levered_beta * us_imp
    print "cost of equity:", round(e_return * 100, 2), "%"

    #wacc calculation
    wacc = e_return * (1 - d_e) + cost_of_debt() * d_e

    return wacc
   

print "wacc:", round(wacc() * 100, 2), "%"



