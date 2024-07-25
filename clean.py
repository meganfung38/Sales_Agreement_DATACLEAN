import pandas as pd


def same_left_right_date(df):
    """populate blank cells in between two cells that have the same sales agreement date
    columns: rc_account_id, brand, mrr, churn_date, EntryCount, dates (incremented by month 8/2019- 6/2024)
    rows: rc_account_id and their sales agreement dates"""
    dates = df.columns[5:]  # date columns start at column F
    for index, row in df.iterrows():  # iterate through account ids
        for date in range(len(dates) - 1):  # iterate through sale agreements for current account id
            left_date = row[dates[date]]  # current sales agreement
            if pd.notna(left_date):  # if current sales agreement is not blank
                for right_date in range(date + 1, len(dates)):  # iterate through sales agreements after current
                    if pd.notna(row[dates[right_date]]):  # found a non-blank cell
                        if left_date == row[dates[right_date]]:  # check if left and right dates are the same
                            # for the current row, populate cells between left and right date
                            for cell in range(date + 1, right_date):  # traverse through dates in between
                                data.iloc[index, data.columns.get_loc(dates[cell])] = left_date
                        break  # stop traversing once we have filled blank cells between two none blank cells
    return df


def fix_date_format(df):
    """changes cell dates and column names from datetime objs to formatted dates"""
    date_columns = data.columns[5:]  # define columns that represent dates
    for col in date_columns:  # iterate through date column cells
        df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%m/%d/%Y')
    # change column names to formatted dates
    formatted_date_columns = pd.to_datetime(date_columns, errors='coerce').strftime('%m/%d/%Y')
    df.rename(columns=dict(zip(date_columns, formatted_date_columns)), inplace=True)
    return df


# calling functions for data clean up
file_input = ''  # define excel export
data = pd.read_excel(file_input)  # open and read excel

data = same_left_right_date(data)  # calling clean up
data = fix_date_format(data)  # format dates

output_file = 'cleaned.xlsx'  # write to this output file
data.to_excel(output_file, index=False)  # writing updated data frame to output file

print("finished clean up")

