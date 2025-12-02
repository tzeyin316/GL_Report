import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime

def first_number(cells):
    for cell in cells:
        if isinstance(cell, str):
            clean = cell.replace("(", "-").replace(")", "").replace(",", "").strip()
            return float(clean)
        elif isinstance(cell, float) and not np.isnan(cell):
            return cell
    return np.nan

def clean_tax(cells):
    for cell in cells:
        if isinstance(cell, str):
            return cell
    return np.nan

def convert_date(value):
    if isinstance(value, datetime):
        return value.date()
    else:
        return value

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()

def output_table(clean_df):
    st.write("Clean Table:")
    st.dataframe(clean_df, width="stretch")

    excel_binary = convert_df_to_excel(clean_df)

    st.download_button(
        label="Download",
        data=excel_binary,
        file_name=f"{uploaded_file.name}_clean.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# main section
st.title("Reformat General Ledger")

category = st.selectbox("Select category:", ["Select a category", "SQL", "AutoCount"])

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

# SQL GL
if category == "SQL" and uploaded_file is not None:
    # Initialize
    final_header = [
    "Date",
    "Reference Code",
    "Account Code",
    "Account Name",
    "Description 1",
    "Desciption 2",
    "Tax",
    "Debit (RM)",
    "Credit (RM)",
    "Ending Balance (RM)",
    ]
    subAccCode, subAccName = "", ""
    final_data = []

    file = pd.ExcelFile(uploaded_file)

    status = st.empty()
    status.success("Start Processing...")  # message

    # Loop through sheets and append to combined_df
    for each_sheet in file.sheet_names:
        df = pd.read_excel(file, sheet_name=each_sheet)
        df = df.drop(df.columns[1], axis=1)
        data_arr = df.values  # convert to numpy array

        header = data_arr[2]
        values = data_arr[3:]
        header_cell = []

        for h in header:
            if isinstance(h, float) and np.isnan(h):  # check if the cell itself is NaN
                header_cell[-1] += 1
            else:
                header_cell.append(1)

        for row in values:
            if isinstance(row[0], str) and row[0].strip().startswith("Code :"):
                subAccCode, subAccName = (
                    row[0].replace("Code :", "").strip().split(" ", 1)
                )
                subAccCode = subAccCode.strip()
                subAccName = subAccName.strip()

            elif isinstance(row[2], float) and np.isnan(row[2]):
                continue

            else:
                # Fixed column positions
                date = row[0]
                ref = row[1]
                desc1 = row[2]
                desc2 = row[3]

                # Calculate start indices based on header_cell
                tax_end = 4 + header_cell[4]
                debit_end = tax_end + header_cell[5]
                credit_end = debit_end + header_cell[6]
                balance_end = credit_end + header_cell[7]

                tax = clean_tax(row[4:tax_end])
                debit = first_number(row[tax_end:debit_end])
                credit = first_number(row[debit_end:credit_end])
                balance = first_number(row[credit_end:balance_end])

                final_data.append(
                    [
                        date,
                        ref,
                        subAccCode,
                        subAccName,
                        desc1,
                        desc2,
                        tax,
                        debit,
                        credit,
                        balance,
                    ]
                )

    clean_df = pd.DataFrame(final_data, columns=final_header)
    clean_df = clean_df.iloc[:-8]

    output_table(clean_df)

    status.success("Done!")  # first message

elif category == "AutoCount" and uploaded_file is not None:
    # Initialize
    subAccCode, subAccName = "", ""
    final_data = []
    # response = ocr(uploaded_file)
    # print(response)

    df = pd.read_excel(uploaded_file, sheet_name=0)

    df = df.drop(df.columns[0], axis=1)
    df = df.iloc[10:]  # drop first 10 rows

    data_arr = df.values  # convert to numpy array

    header = data_arr[4]
    values = data_arr[0:]
    h1_cell = []
    h2_cell = []

    desc_index = np.where(header == "Description")[0][0]
    # Split the array
    h1 = header[:desc_index+1]  # include "Description"
    h2 = header[desc_index+1:]  # after "Description"

    for h in h1:
        if isinstance(h, float) and np.isnan(h):  # check if the cell itself is NaN
            h1_cell[-1] += 1
        else:
            h1_cell.append(1)

    for h in reversed(h2):
        if isinstance(h, float) and np.isnan(h):  # check if the cell itself is NaN
            h2_cell[-1] += 1
        else:
            h2_cell.append(1)

    header_cell = h1_cell + list(reversed(h2_cell))

    # variables define
    i = 0
    ref_index = sum(header_cell[:2])
    desc_index = sum(header_cell[:3])
    debit_start = sum(header_cell[:4]) + 1
    debit_end    = debit_start + header_cell[4]
    credit_end   = debit_end + header_cell[5]
    balance_end  = credit_end + header_cell[6]

    while i < len(values):
        row = values[i]   
        print(i)

        if isinstance(row[0], str) and row[0].strip().startswith("Account Code:"):
            account = [x for x in row if not (isinstance(x, float) and np.isnan(x))]
            subAccCode = account[1]
            subAccName = account[2]
            i = i+1     

        elif (isinstance(row[ref_index], float) and np.isnan(row[ref_index])) or (isinstance(row[0], str) and row[0].strip().startswith("Date")):
            i = i+1     
            continue

        else:
            # Fixed column positions
            date   = convert_date(row[0])
            journal = row[header_cell[0]]
            ref1 = row[ref_index]
            desc1 = row[desc_index]

            # Calculate start indices based on header_cell
            debit   = first_number(row[debit_start:debit_end])
            credit  = first_number(row[debit_end:credit_end])
            balance = first_number(row[credit_end:balance_end])
            
            if isinstance(desc1, float) and np.isnan(desc1):
                ref2, desc2 = np.nan, np.nan
                i = i+1
            else:
                second_row = values[i+2] if i+2 < len(values) else None
                ref2 = second_row[ref_index]
                third_row = values[i+3] if i+3 < len(values) else None
                desc2 = third_row[desc_index]
                i = i+4

            print(date, subAccCode, subAccName, journal, ref1, ref2, desc1, desc2, debit, credit, balance)
            final_data.append([date, subAccCode, subAccName, journal, ref1, ref2, desc1, desc2, debit, credit, balance])


    final_header = [
        "Date",
        "Account Code",
        "Account Name",
        "Journal",
        "Reference 1",
        "Reference 2",
        "Description 1",
        "Description 2",
        "Debit (RM)",
        "Credit (RM)",
        "Ending Balance (RM)",
    ]
    clean_df = pd.DataFrame(final_data, columns=final_header)
    output_table(clean_df)
