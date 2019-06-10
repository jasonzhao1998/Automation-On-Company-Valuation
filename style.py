from helper import *
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font


def style_range(ws, start, end, fill=PatternFill(), font=Font(), border=Border(),
                alignment=Alignment()):
    letter1, num1 = start[0], start[1:]
    letter2, num2 = end[0], end[1:]
    if letter1 == letter2:  # column
        for i in range(int(num1), int(num2) + 1):
            ws[letter1 + str(i)].font = font
            ws[letter1 + str(i)].fill = fill
            ws[letter1 + str(i)].border = border
            ws[letter1 + str(i)].alignment = alignment
    elif num1 == num2:  # row
        for i in range(ord(letter2) - ord(letter1) + 1):
            ws[chr(ord(letter1) + i) + num1].font = font
            ws[chr(ord(letter1) + i) + num1].fill = fill
            ws[chr(ord(letter1) + i) + num1].border = border
            ws[chr(ord(letter1) + i) + num1].alignment = alignment
    else:
        print("ERROR: style_range")
        exit(1)


def style_ws(ws, sheet_name, is_df, bs_df, cf_df, fye):
    border = Side(border_style="thin", color="000000")

    # Insert empty column to beginning
    ws.insert_cols(1)

    letter, num = ws.dimensions.split(':')[1][0], ws.dimensions.split(':')[1][1:]

    ws.sheet_view.showGridLines = False  # No grid lines
    ws.move_range('C1:' + letter + '1', rows=4)  # Move year row down
    ws.column_dimensions['B'].width = 50  # Change width of labels

    ws['B2'] = sheet_name
    ws['B2'].font = Font(bold=True)
    ws['B2'].fill = PatternFill(fill_type='solid', fgColor='bababa')
    ws['B3'] = "($ in millions of U.S. Dollar)"
    ws['B3'].font = Font(italic=True)
    style_range(ws, 'B3', letter + '3', fill=PatternFill(fill_type='solid', fgColor='bababa'))

    # Central element Annual
    ws[chr((ord('C') + ord(letter)) // 2) + '3'] = "Annual"
    ws[chr((ord('C') + ord(letter)) // 2) + '3'].font = Font(bold=True)
    ws[chr((ord('C') + ord(letter)) // 2) + '4'] = "FYE " + fye

    # Year row style
    style_range(ws, 'C5', letter + '5', font=Font(bold=True, underline="single"),
                border=Border(top=border, bottom=border),
                alignment=Alignment(horizontal="center", vertical="center"))

    # Label column
    style_range(ws, 'B7', 'B' + num, fill=PatternFill(fill_type='solid', fgColor='dddddd'))

    # Style sum rows
    for cell in [letter + str(i + 1) for i in range(int(num) - 1)]:
        if type(ws[cell].value) is str and 'SUM' in ws[cell].value and len(ws[cell].value) < 30:
            num = cell[1:]
            ws['B' + num].font = Font(bold=True)
            style_range(ws, 'C' + num, letter + num, font=Font(bold=True),
                        border=Border(top=border))

    # Style specific rows
    def style_row(ws, label, sheet_name, is_df, bs_df, cf_df, border_bool=True, bold_bool=True,
                  underline=None):
        df = None
        num = 0
        if sheet_name == "Income Statement":
            num = str(int(excel_cell(is_df, searched_label(is_df.index, label),
                                     is_df.columns[0])[1:]))
        elif sheet_name == "Balance Sheet":
            num = str(int(excel_cell(bs_df, searched_label(bs_df.index, label),
                                     bs_df.columns[0])[1:]))
        elif sheet_name == "Cashflow Statement":
            num = str(int(excel_cell(cf_df, searched_label(cf_df.index, label),
                                     cf_df.columns[0])[1:]))
        ws['B' + num].font = Font(bold=True, underline=underline)
        border_style = Border(top=border) if border_bool else Border()
        bold_style = Font(bold=True) if bold_bool else Font()
        style_range(ws, 'C' + num, letter + num, font=bold_style, border=border_style)

    if sheet_name == "Income Statement":
        style_row(ws, "total sales", sheet_name, is_df, bs_df, cf_df, False)
        style_row(ws, "gross income", sheet_name, is_df, bs_df, cf_df)
        style_row(ws, "ebit operating income", sheet_name, is_df, bs_df, cf_df)
        style_row(ws, "ebit operating income", sheet_name, is_df, bs_df, cf_df)
        style_row(ws, "pretax income", sheet_name, is_df, bs_df, cf_df)
        style_row(ws, "net income", sheet_name, is_df, bs_df, cf_df)
        style_row(ws, "driver ratio", sheet_name, is_df, bs_df, cf_df, underline="single",
                  border_bool=False)
    elif sheet_name == "Balance Sheet":
        style_row(ws, "total shareholder equity", sheet_name, is_df, bs_df, cf_df, bold_bool=False,
                  border_bool=False)
        style_row(ws, "total liabilit shareholder equity", sheet_name, is_df, bs_df, cf_df,
                  border_bool=False)
        style_row(ws, "driver ratio", sheet_name, is_df, bs_df, cf_df, underline="single",
                  border_bool=False)
    elif sheet_name == "Cashflow Statement":
        style_row(ws, "net operating cash flow cf", sheet_name, is_df, bs_df, cf_df)
        style_row(ws, "cash balance", sheet_name, is_df, bs_df, cf_df, border_bool=False)
        style_row(ws, "driver ratio", sheet_name, is_df, bs_df, cf_df, underline="single",
                  border_bool=False)
