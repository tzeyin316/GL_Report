import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, timezone


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

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()

def output_table(clean_df):
    st.write("Clean Table:")
    st.dataframe(clean_df)

    excel_binary = convert_df_to_excel(clean_df)

    st.download_button(
        label="Download",
        data=excel_binary,
        file_name=f"{uploaded_file.name}_clean.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.title("SQL GL Report")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
status = st.empty()

if uploaded_file is not None:

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

    my_bar = st.progress(0, text="Processing...")

    # Loop through sheets and append to combined_df
    for idx, each_sheet in enumerate(file.sheet_names):
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

        my_bar.progress(int((idx)/len(file.sheet_names)*100), text="Processing...")

    clean_df = pd.DataFrame(final_data, columns=final_header)
    clean_df = clean_df.iloc[:-8]

    output_table(clean_df)
    my_bar.empty()