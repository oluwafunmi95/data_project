import pandas as pd

datafile = "2013_ERCOT_Hourly_Load_Data.xls"

def parse_file(datafile):
    df = pd.read_excel(datafile, sheet_name=0)

    data = df.values.tolist()

    print("\nList Comprehension")
    print("data[3][2]:", data[3][2])

    print("\nCells in a nested loop:")    
    for row in range(df.shape[0]):
        for col in range(df.shape[1]):
            if row == 50:
                print(df.iat[row, col], end=' ')

    ### other useful methods:
    print("\nROWS, COLUMNS, and CELLS:")
    print("Number of rows in the sheet:", df.shape[0])
    print("Type of data in cell (row 3, col 2):", type(df.iat[3, 2]))
    print("Value in cell (row 3, col 2):", df.iat[3, 2])
    print("Get a slice of values in column 3, from rows 1-3:")
    print(df.iloc[1:4, 3].tolist())

    print("\nDATES:")
    print("Type of data in cell (row 1, col 0):", type(df.iat[1, 0]))
    exceltime = df.iat[1, 0]
    print("Time in Excel format:", exceltime)
    print("Convert time to a Python datetime:", pd.to_datetime(exceltime))

parse_file(datafile)
