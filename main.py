<<<<<<< HEAD
import os
import numpy as np
import pandas as pd

FILENAME = "Nvidia.xlsx"


def main():
    df = pd.read_excel('asset/' + FILENAME, header=4, index_col=0)
    print(df)
    print(df.dtypes)
    print(df.index)
    print(df.columns)

if __name__ == "__main__":
    main()
