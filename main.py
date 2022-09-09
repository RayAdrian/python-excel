import pandas as pd
import openpyxl

def read_excel():
    # Read Sheet1 of the excel file
    df_one = pd.read_excel('test_excel.xlsx', sheet_name="Sheet1")
    print("Printing excel contents\n")
    print(df_one)

    # Read Sheet2 of the excel file
    df_two = pd.read_excel('test_excel.xlsx', sheet_name="Sheet2")
    print("\n\nPrinting excel Sheet2 contents")
    print(df_two)


def write_excel():
    # Read Sheet1 of excel file
    df_one = pd.read_excel('test_excel.xlsx', sheet_name="Sheet1").to_dict(orient="records")
    print("Old values\n\n")
    print(df_one)

    # Change value of the first name
    df_one[0]["Name"] = "Test"

    # Convert back to dataframe before writing to excel
    df_final = pd.DataFrame(df_one)

    # Update the excel sheet
    with pd.ExcelWriter('test_excel.xlsx', mode='a', if_sheet_exists="replace") as writer:
        df_final.to_excel(writer, sheet_name="Sheet1", index=False)

    # Read the updated one
    df_one = pd.read_excel('test_excel.xlsx', sheet_name="Sheet1").to_dict(orient="records")
    print("New values\n\n")
    print(df_one)


if __name__ == "__main__":
#    read_excel()
   write_excel()