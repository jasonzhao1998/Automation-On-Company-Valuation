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

def capm():
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

    #beta calculation
    beta = pd.read_excel('asset/betas.xlsx', header=9, index_col=0, na_filter=False)
    unlevered_beta = beta.loc['Semiconductor Equip', 'Unlevered beta corrected for cash']  # should be extracted from beta.xlsx
    d_e = 0  # should be calculated from company balance sheet
    levered_beta = unlevered_beta * (1 + (1 - t_m) * d_e)

    #capm calculation
    e_return = us_10y_treasury_bond + levered_beta * us_imp

    return e_return

print capm()



