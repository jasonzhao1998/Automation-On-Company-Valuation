"""Implementation for rendering excel output's style."""
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from helper import excel_cell, searched_label

IS = "Income Statement"
BS = "Balance Sheet"
CF = "Cashflow Statement"


def style_range(ws, start, end, fill=None, font=None, border=None,
                alignment=None, percentage=False, currency=False):
    """Change excel style for a column or row."""
    letter1, num1 = start[0], start[1:]
    letter2, num2 = end[0], end[1:]
    if num1 == num2:  # row
        for i in range(ord(letter2) - ord(letter1) + 1):
            if font:
                ws[chr(ord(letter1) + i) + num1].font = font
            if fill:
                ws[chr(ord(letter1) + i) + num1].fill = fill
            if border:
                ws[chr(ord(letter1) + i) + num1].border = border
            if alignment:
                ws[chr(ord(letter1) + i) + num1].alignment = alignment
            if currency:
                ws[chr(ord(letter1) + i) + num1].number_format = '$#,##'
            elif percentage and isinstance(ws[chr(ord(letter1) + i) + num1].value, str):
                if '365' in ws[chr(ord(letter1) + i) + num1].value:
                    return
                elif '366' in ws[chr(ord(letter1) + i) + num1].value:
                    return
                elif '+' in ws[chr(ord(letter1) + i) + num1].value:
                    return
                ws[chr(ord(letter1) + i) + num1].number_format = '0.00%'
    elif letter1 == letter2:  # column
        for i in range(int(num1), int(num2) + 1):
            if font:
                ws[letter1 + str(i)].font = font
            if fill:
                ws[letter1 + str(i)].fill = fill
            if border:
                ws[letter1 + str(i)].border = border
            if alignment:
                ws[letter1 + str(i)].alignment = alignment
    else:
        print("ERROR: style_range")
        exit(1)


def style_ws(ws, sheet_name, is_df, bs_df, cf_df, fye, unit):
    """Change excel style for a worksheet."""
    if sheet_name == "Income Statement":
        cur_df = is_df
    elif sheet_name == "Balance Sheet":
        cur_df = bs_df
    elif sheet_name == "Cashflow Statement":
        cur_df = cf_df

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
    if unit == 'm':
        ws['B3'] = "($ in millions of U.S. Dollar)"
    else:
        ws['B3'] = "($ in billions of U.S. Dollar)"
    ws['B3'].font = Font(italic=True)
    style_range(ws, 'B3', letter + '3', fill=PatternFill(fill_type='solid', fgColor='bababa'))

    # Central element Annual
    ws[chr((ord('C') + ord(letter)) // 2) + '3'] = "Annual"
    ws[chr((ord('C') + ord(letter)) // 2) + '3'].font = Font(bold=True)
    ws[chr((ord('C') + ord(letter)) // 2) + '4'] = "FYE " + fye
    # Center across selection
    ws[chr((ord('C') + ord(letter)) // 2) + '3'].alignment = Alignment(horizontal='centerContinuous')
    ws[chr((ord('C') + ord(letter) + 1) // 2) + '3'].alignment = Alignment(horizontal='centerContinuous')
    ws[chr((ord('C') + ord(letter)) // 2) + '4'].alignment = Alignment(horizontal='centerContinuous')
    ws[chr((ord('C') + ord(letter) + 1) // 2) + '4'].alignment = Alignment(horizontal='centerContinuous')

    # Year row style
    style_range(ws, 'C5', letter + '5', font=Font(bold=True, underline="single"),
                border=Border(top=border, bottom=border),
                alignment=Alignment(horizontal="center", vertical="center"))

    # Label column
    style_range(ws, 'B7', 'B' + num, fill=PatternFill(fill_type='solid', fgColor='dddddd'))

    # Style sum rows
    for cell in [letter + str(i + 1) for i in range(int(num) - 1)]:
        if isinstance(ws[cell].value, str) and 'SUM' in ws[cell].value and len(ws[cell].value) < 30:
            num = cell[1:]
            ws['B' + num].font = Font(bold=True)
            style_range(ws, 'C' + num, letter + num, font=Font(bold=True),
                        border=Border(top=border), currency=True)

    # Style specific rows
    def style_row(ws, label, df, border_bool=True, bold_bool=True,
                  underline=None, currency=False):
        num = str(int(excel_cell(df, searched_label(df.index, label), df.columns[0])[1:]))
        ws['B' + num].font = Font(bold=True, underline=underline)
        border_style = Border(top=border) if border_bool else Border()
        bold_style = Font(bold=True) if bold_bool else Font()
        style_range(ws, 'C' + num, letter + num, font=bold_style, border=border_style,
                    currency=currency)

    if sheet_name == "Income Statement":
        style_row(ws, "total sales", cur_df, False, currency=True)
        style_row(ws, "gross income", cur_df, currency=True)
        style_row(ws, "ebit operating income", cur_df, currency=True)
        style_row(ws, "pretax income", cur_df, currency=True)
        style_row(ws, "net income", cur_df, currency=True)
    elif sheet_name == "Balance Sheet":
        style_row(ws, "total shareholder equity", cur_df, bold_bool=False, border_bool=False,
                  currency=True)
        style_row(ws, "total liabilit shareholder equity", cur_df, border_bool=False)
    elif sheet_name == "Cashflow Statement":
        style_row(ws, "net operating cash flow cf", cur_df, currency=True)
        style_row(ws, "cash balance", cur_df, border_bool=False, currency=True)
    style_row(ws, "driver ratio", cur_df, underline="single", border_bool=False)
    driver_df = cur_df.loc["Driver Ratios":]

    # Driver ratios style
    for ratio in driver_df[driver_df.index.notna()].iloc[1:].index:
        start = excel_cell(cur_df, ratio, cur_df.columns[0])
        end = letter + str(int(start[1:]))
        style_range(ws, start, end, percentage=True)


def write_n_style_summary(ws, is_df, bs_df, cf_df, fye, is_unit, years):
    """Note that number of years is fixed here."""
    border = Side(border_style="thin", color="000000")
    ws.sheet_view.showGridLines = False  # No grid lines
    ws.column_dimensions['B'].width = 30  # Change width of labels

    # Header
    ws['B2'] = "Financial Overview"
    ws['B2'].font = Font(bold=True)
    ws['B2'].fill = PatternFill(fill_type='solid', fgColor='bababa')
    if is_unit == 'm':
        ws['B3'] = "($ in millions of U.S. Dollar)"
    else:
        ws['B3'] = "($ in billions of U.S. Dollar)"
    ws['B3'].font = Font(italic=True)
    ws['B3'].fill = PatternFill(fill_type='solid', fgColor='bababa')

    # Summary Financials
    ws.column_dimensions['C'].width = 15
    ws['C5'] = "Summary Financials"
    style_range(ws, 'C5', 'I5', fill=PatternFill(fill_type='solid', fgColor='bababa'),
                font=Font(bold=True), alignment=Alignment(horizontal="centerContinuous"))
    ws['C6'] = "FYE " + fye
    style_range(ws, 'C6', 'I6', alignment=Alignment(horizontal="centerContinuous"),
                border=Border(top=border, bottom=border))
    for i in range(len(years)):
        ws[chr(ord('D') + i) + '7'] = years[i]
    style_range(ws, 'D7', 'I7', font=Font(bold=True, underline="single"), alignment=Alignment(horizontal="center"))
    ws['C8'], ws['C9'] = "Revenue", "Growth %"
    ws['C11'], ws['C12'], ws['C13'] = "Gross Profit", "Margin %", "Growth %"
    ws['C15'], ws['C16'], ws['C17'] = "EBITDA", "Margin %", "Growth %"
    ws['C19'],  ws['C20'] = "EPS", "Growth %"
    ws['C22'], ws['C23'] = "ROA", "ROE"
    for i in range(len(years)):
        revenue = excel_cell(is_df, searched_label(is_df.index, "total sales"), years[i])
        prev_revenue = chr(ord(revenue[0]) - 1) + revenue[1:]
        ws[chr(ord('D') + i) + '8'] = "='{}'!{}".format(IS, revenue)
        ws[chr(ord('D') + i) + '9'] = ws[chr(ord('D') + i) + '8'].value + "/'{}'!{}-1".format(
            IS, prev_revenue
        )
        gross_profit = excel_cell(is_df, searched_label(is_df.index, "gross income"), years[i])
        prev_gross_profit = chr(ord(gross_profit[0]) - 1) + gross_profit[1:]
        ws[chr(ord('D') + i) + '11'] = "='{}'!{}".format(IS, gross_profit)
        ws[chr(ord('D') + i) + '12'] = '=' + chr(ord('D') + i) + '11/' + chr(ord('D') + i) + '8'
        ws[chr(ord('D') + i) + '13'] = ws[chr(ord('D') + i) + '11'].value  + "/'{}'!{} - 1".format(
            IS, prev_gross_profit
        )
        ebitda = excel_cell(is_df, searched_label(is_df.index, "ebitda"), years[i])
        prev_ebitda = chr(ord(ebitda[0]) - 1) + ebitda[1:]
        ws[chr(ord('D') + i) + '15'] = "='{}'!{}".format(IS, ebitda)
        ws[chr(ord('D') + i) + '16'] = '=' + chr(ord('D') + i) + '15/' + chr(ord('D') + i) + '8'
        ws[chr(ord('D') + i) + '17'] = ws[chr(ord('D') + i) + '15'].value  + "/'{}'!{} - 1".format(
            IS, prev_ebitda
        )
        eps = excel_cell(is_df, searched_label(is_df.index, "eps diluted"), years[i])
        prev_eps = chr(ord(eps[0]) - 1) + eps[1:]
        ws[chr(ord('D') + i) + '19'] = "='{}'!{}".format(IS, eps)
        ws[chr(ord('D') + i) + '20'] = ws[chr(ord('D') + i) + '19'].value  + "/'{}'!{} - 1".format(
            IS, prev_eps
        )
        ws[chr(ord('D') + i) + '22'] = years[i]
        ws[chr(ord('D') + i) + '23'] = years[i]
    style_range(ws, 'C23', 'I23', border=Border(bottom=border))
    style_range(ws, 'C6', 'C23', border=Border(left=border))
    style_range(ws, 'I6', 'I23', border=Border(right=border))
    ws['C23'].border = Border(left=border, bottom=border)
    ws['I23'].border = Border(right=border, bottom=border)
    ws['C6'].border = Border(left=border, bottom=border, top=border)
    ws['I6'].border = Border(right=border, bottom=border, top=border)