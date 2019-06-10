import pandas as pd
import numpy as np


ROUNDING_DIGIT = 4


def excel_cell(df, row_label, col_label, nearby_label=None):
    """Returns corresponding excel cell position given row label and column label. 
    Note that if there are more than 26 columns, this function does not work properly."""
    if not row_label:
        print("ERROR: excel_cell")
        exit(1)
    letter = chr(ord('A') + df.columns.get_loc(col_label) + 2)
    row_mask = df.index.get_loc(row_label)
    if type(row_mask) is int:
        return letter + str(3 + row_mask)
    else:
        nearby_index = df.index.get_loc(nearby_label)
        matched_indices = [i for i, j in enumerate(row_mask) if j]
        distance_vals = [abs(nearby_index - i) for i in matched_indices]
        return letter + str(3 + matched_indices[distance_vals.index(min(distance_vals))])


def searched_label(labels, target):
    """Returns target label from a list of DataFrame labels."""
    score_dict = {label: 0 for label in labels}

    for word in target.split():
        for label in set(labels):
            if word.lower() in str(label).lower():
                score_dict[label] += 1

    if sum(score_dict.values()) == 0:
        return ""
    def compare(pair):
        if type(pair[0]) is str:
            return len(pair[0])
        return 0
    return max(sorted(score_dict.items(), key=compare), key=lambda pair: pair[1])[0]


def empty_unmodified(df, yrs_to_predict):
    unmodified = df.iloc[:, -yrs_to_predict] == '0'
    df.loc[unmodified, :] = np.nan
    df.index = [i if i not in list(df.index[unmodified]) else np.nan for i in list(df.index)]


def initialize_ratio_row(df, top_label, bot_label, new_label, nearby_label=None):
    """Create a new label and set a fractional formula for initialization."""
    df.loc[new_label] = [
        '={}/{}'.format(excel_cell(df, top_label, col, nearby_label),
                        excel_cell(df, bot_label, col))
        for col in df.columns
    ]


def insert_before(df, new_df, label):
    """Insert new DataFrame before the corresponding label row."""
    index = list(df.index).index(searched_label(df.index, label))
    return pd.concat([df.iloc[:index], new_df, df[index:]])


def insert_after(df, new_df, label):
    """Insert new DataFrame after the corresponding label row."""
    index = list(df.index).index(searched_label(df.index, label))
    return pd.concat([df.iloc[:index + 1], new_df, df[index + 1:]])


def add_empty_row(df):
    """Adds an empty row to the bottom of DataFrame."""
    df.loc["null"] = np.nan
    df.index = list(df.index)[:-1] + [np.nan]


def add_yr_column(df):
    """Appends one empty column representing year into DataFrame."""
    cur_yr = str(df.columns[len(df.columns) - 1])
    if cur_yr[-1] == 'E':
        cur_yr = str(int(cur_yr[:-1]) + 1) + 'E'
    else:
        cur_yr = str(int(cur_yr) + 1) + 'E'
    array = ['0' if i else np.nan for i in df.iloc[:,-1].notna().values]
    df.insert(len(df.columns), cur_yr, array)


def add_growth_rate_row(df, label, new_label):
    df.loc[new_label] = [np.nan] + [
        '={}/{}-1'.format(
            excel_cell(df, label, df.columns[i + 1]), excel_cell(df, label, df.columns[i])
        ) for i in range(len(df.columns) - 1)
    ]


def driver_extend(df, row_label, how, last_given_yr, yrs_to_predict, num_excluded=0):
    if how is "round":
        formula = "=ROUND(" + excel_cell(df, row_label, last_given_yr) + ',' + \
                  str(ROUNDING_DIGIT) + ')'
    elif how is "avg":
        formula = "=AVERAGE(" + excel_cell(df, row_label, df.columns[0 + num_excluded]) + ':' + \
                  excel_cell(df, row_label, last_given_yr) + ')'
    df.loc[row_label].iloc[-yrs_to_predict] = formula
    temp = excel_cell(df, row_label, df.columns[-yrs_to_predict])
    df.loc[row_label].iloc[-yrs_to_predict + 1:] = '=' + temp


def fixed_extend(df, row_label, how, yrs):
    """Predicts the corresponding row of data only using data from current row."""
    if not row_label:
        print("Empty row_label in fixed_extend")
        return

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
        df.at[row_label, df.columns[-yrs:]] = 0
    else:
        print("ERROR: fixed_extend")
        exit(1)


def sum_formula(df, row_label, col_label, start_label=None, offset=0):
    end_label = df.loc[:row_label].index[-2]

    if start_label:
        start_label = df.index[df.index.get_loc(start_label) + offset]
    else:
        for i in range(df.index.get_loc(row_label), 0, -1):
            if pd.isna(df.loc[:, col_label].iloc[i]):
                start_label = df.index[i + 1]
                break
    formula = 'SUM({}:{})'.format(
        excel_cell(df, start_label, col_label, row_label), excel_cell(df, end_label, col_label, row_label)
    )
    return formula
