import pandas as pd
import xlwt

EXCEL_FILE_NAME = "Referrals.xlsx"
FIRST_SHEET_NAME = "Work Received inc amavat 2021"
SECOND_SHEET_NAME = "Work Referred inc amavat 2021"
OUTPUT = "Output/"


def defaultExcel():
    excel = pd.ExcelFile("Files/" + EXCEL_FILE_NAME)
    df1 = excel.parse(FIRST_SHEET_NAME)
    df2 = excel.parse(SECOND_SHEET_NAME)
    df = [df1, df2]
    return df


def getCompanyName(df):
    companies = []
    for index in range(len(df["To Company"])):
        if index > 0:
            if df["To Company"][index] != df["To Company"][index-1] :
                if "Total" not in str(df["To Company"][index]):
                    companies.append(df["To Company"][index])
                    # print(df["To Company"][index])

    return companies

def createSheet(df, file_name):
    # workbook = xlwt.Workbook()
    # sheet = workbook.add_sheet("Sheet1")

    # workbook.save(FILE_NAME)
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    # Write data to an excel
    df[0].to_excel(writer,sheet_name=FIRST_SHEET_NAME,index=False)
    df[0].to_excel(writer,sheet_name=SECOND_SHEET_NAME,index=False)
    # Get sheet for conditional formatting 
    # worksheet = writer.sheets['Sheet1']
    # Add conditional formatting for Age column
    # worksheet.conditional_format('B2:B5', {'type': '2_color_scale'})
    # Close writer
    writer.close()

def createSheets():
    df = defaultExcel()
    companies = getCompanyName(df[0])
    # for i in range(len(companies)):
    for i in range(4):
        array = [companies[i], companies[i]+" Total"]
        new_df = df[0].loc[df['To Company'].isin(array)]
        print("Creating", companies[i]+".xlsx")
        createSheet(new_df, OUTPUT+companies[i]+".xlsx")
        print("Done")

def main():
    createSheets()
    
    
    print("Extracting")
    print("Finish")

main()

#pecah df jadi 2
#ambil nama companies 
#tulis ke excel, check kalo sudah ada filenya