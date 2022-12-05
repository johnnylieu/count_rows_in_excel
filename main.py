import xlrd as xl
import pandas as pd
from openpyxl import load_workbook

files = ['finalized_S5cz2H2ey8Em6PPkun6FsM_aggregate_2022-12-01_scandit_localization_BJs_2022-12-01_localization_aggregate.xlsx']

for file in files:
    count = 0

    wrkbk = load_workbook(file)
    sh = wrkbk.active
    total_rows = sh.max_row

    aisle_value = sh.cell(row=2, column=7).value

    for i in range(1, total_rows+1):
        if (sh.cell(row = i, column=7).value) != "":
            count+=1
    print(f'\n🖩Sup Playa, you loaded {len(files)} localization report & the count is: {count}.\n\n🤖Stay Positive and keep grindin big dawg, we finna get there soon!')