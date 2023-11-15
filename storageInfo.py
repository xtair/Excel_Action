import excelActions
from excelActions import *

# could be change when address changes
DIR = '/Users/xuanzhentai/Happy Distro Files/Happy Distro Excel/Sample File/'
satFileName = 'Weekly Inventory Allocation Form_10.28_v4周末分货.xlsx'
wedFileName = 'Weekly Inventory Allocation Form_11.02_V3周三加分.xlsx'
addFileName = 'Week 10.29~11.4 Additional.xlsx'
soFileName = 'Sales_report.xlsx'
salesFileName = 'Copy of List of All Sales_by 10302023.xlsx'
NameList = ['vicky', 'andy', 'grace', 'aiden', 'david', 'davey', 'shirley', 'kiki', 'sam', 'skyler', 'jason']


def headerLowerCase(dfOne, mult):
    if mult:
        for key, df in dfOne.items():
            df.columns = [str(col).lower() for col in df.columns]
    else:
        dfOne.columns = [str(col).lower() for col in dfOne.columns]


def readfile():
    # read file
    dfSat = dfWed = dfAdd = dfSO = dfSales = None
    try:
        dfSat = pd.read_excel(DIR + satFileName, sheet_name=None, skiprows=1, engine='openpyxl')
        dfWed = pd.read_excel(DIR + wedFileName, sheet_name=None, skiprows=1, engine='openpyxl')
        dfAdd = pd.read_excel(DIR + addFileName, engine='openpyxl')
        dfSO = pd.read_excel(DIR + soFileName, engine='openpyxl')
        dfSales = pd.read_excel(DIR + salesFileName, engine='openpyxl')
    except Exception as e:
        # 处理所有类型的异常
        print(f"发生了异常: {e}")
        print("请确认文件地址已经文件名称是否正确")

    # header to lower case
    headerLowerCase(dfSat, True)
    headerLowerCase(dfWed, True)
    headerLowerCase(dfAdd, False)
    headerLowerCase(dfSO, False)
    headerLowerCase(dfSales, False)

    return [dfSat, dfWed, dfAdd, dfSO, dfSales]


def getSum():
    """
    This method read target sheets of files, and get sum of the same sku base on group leader.
    :return: None
    """
    read_file = readfile()
    dfSat = read_file[0]
    dfWed = read_file[1]
    dfAdd = read_file[2]
    dfSO = read_file[3]
    dfSales = read_file[4]

    # 获取工作表名称
    sat_sheets = set(dfSat.keys())
    wed_sheets = set(dfWed.keys())

    # 确定缺失的工作表
    missing_in_sat = wed_sheets - sat_sheets
    missing_in_wed = sat_sheets - wed_sheets

    # Dictionary to store the processed dataframes
    processed_dfs = getAllSheetName(dfSat)

    rest(processed_dfs, dfSat, dfWed, dfAdd, dfSO, dfSales)

    print("All Done!")


def getAllSheetName(df):     # need to be changed
    # Dictionary to store the processed dataframes
    processed_dfs = {}
    # Loop through all sheets
    index = 0
    for sheet_name, df in df.items():
        # Store the processed dataframe in the dictionary
        if 3 < index < 16:
            processed_dfs[sheet_name] = df
            # Optionally print to track progress
            print(f"Processed sheet: {sheet_name}")
        index += 1
    return processed_dfs


def rest(processed_dfs, dfSat, dfWed, dfAdd, dfSO, dfSales):
    for name in ['vicky']:  # no Kiki
        count = 0
        fileName = ''
        # loop each sheets in both Excel file
        for ele in processed_dfs:
            dfSat_sheet = None
            dfWed_sheet = None
            dfs = []
            if ele in dfSat:
                dfSat_sheet = dfSat[ele]
                dfSat_sheet = renameDuplicated(dfSat_sheet)
            if ele in dfWed:
                dfWed_sheet = dfWed[ele]
                dfWed_sheet = renameDuplicated(dfWed_sheet)
            if dfSat_sheet and dfWed_sheet:
                dfWed_sheet = addMissingColumn(dfSat_sheet, dfWed_sheet)
                # print(f"Wed: {dfWed_sheet.columns}")
                dfSat_sheet = addMissingColumn(dfWed_sheet, dfSat_sheet)
                # print(f"Sat: {dfSat_sheet.columns}")

            # List of all DataFrames you want to concatenate, 0 means show sku and flavor, 1 means hide sku and flavor
            dfs.append(mergeAction(dfSat_sheet, dfWed_sheet, dfAdd, dfSO, dfSales, name, "ca", 0))
            # Create an empty DataFrame with one column
            dfs.append(pd.DataFrame({'': [None] * len(dfSat_sheet)}))
            dfs.append(mergeAction(dfSat_sheet, dfWed_sheet, dfAdd, dfSO, dfSales, name + ".1", "tx", 1))
            # Concatenate all DataFrames in the list horizontally
            combined_df = pd.concat(dfs, axis=1)

            if count == 0:
                fileName = createNewFile(combined_df, ele, name)
            else:
                addNewSheet(fileName, combined_df, ele)
            count += 1
        colorFileAndInsertCol(excelActions.DIR + fileName + 'Sum.xlsx')
        print(name + 'Sum file done!')

