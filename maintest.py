import pandas as pd
import xlwt

EXCEL_FILE_NAME = "Referrals.xlsx"
FIRST_SHEET_NAME = "Work Received inc amavat 2021"
SECOND_SHEET_NAME = "Work Referred inc amavat 2021"
OUTPUT = "Output/"



df = pd.read_excel('Files/Referrals.xlsx', sheet_name=[FIRST_SHEET_NAME, SECOND_SHEET_NAME])
df1 = df[FIRST_SHEET_NAME]
df2 = df[SECOND_SHEET_NAME]

companies_sheet_1 = []
to_company = "To Company"
for index in range(len(df1[to_company])):
    if index > 0:
        if df1[to_company][index] != df1[to_company][index-1] :
            if "Total" not in str(df1[to_company][index]):
                companies_sheet_1.append(df1[to_company][index])

companies_sheet_2 = []
from_company = "From Company"
for index in range(len(df2[from_company])):
    if index > 0:
        if df2[from_company][index] != df2[from_company][index-1] :
            if "Total" not in str(df2[from_company][index]):
                companies_sheet_2.append(df2[from_company][index])

# print(companies_sheet_1)
# print(companies_sheet_2)

for i in range(4):
    # array_1 = [companies_sheet_1[i], companies_sheet_1[i]+" Total"]
    new_df_1 = df1.loc[df1[to_company].isin([companies_sheet_1[i], companies_sheet_1[i]+" Total"])]
    new_df_2 = df2.loc[df2[to_company].isin([companies_sheet_2[i], companies_sheet_2[i]+" Total"])]
    dfs = [new_df_2, new_df_2]
    print("Creating", companies_sheet_1[i]+".xlsx")
    writer = pd.ExcelWriter(companies_sheet_1[i]+".xlsx", engine='xlsxwriter')

    for sheet_name in dfs:
        dfs[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
    # createSheet(new_df, OUTPUT+companies[i]+".xlsx")
    print("Done")