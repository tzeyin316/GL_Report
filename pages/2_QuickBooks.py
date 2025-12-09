import streamlit as st
import requests
import base64
import pandas as pd
from io import BytesIO
from datetime import datetime, timezone

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()    
        
def extract_values(row):
    all_values = []

    # Data row with ColData
    if "ColData" in row and row.get("type") == "Data":
        all_values.append([col.get("value", "") for col in row["ColData"]])

    # Nested rows
    if "Rows" in row and "Row" in row["Rows"]:
        for subrow in row["Rows"]["Row"]:
            all_values.extend(extract_values(subrow))  # recursion

    return all_values


st.title("QuickBooks GL Report")

start_date = st.date_input("Start date")
end_date = st.date_input("End date")

if st.button("Run workflow"):
    with st.spinner("Running workflow..."):
        webhook = "https://hook.us2.make.com/yv6elkcn4q7f6yf9zhzruscwy4x2eau1"  # your webhook

        # Send payload to Make
        response = requests.post(
            webhook,
            json={
                "value": "qb",
                "start_date": str(start_date),
                "end_date": str(end_date)
            }
        )

        # Try JSON first — if fails, show text
        try:
            data = response.json()
            # st.write(data)

            # header - column titles 
            columns = data["Columns"]["Column"]
            headers = [col["ColTitle"] for col in columns]

            # rows - values
            all_values = []
            for row in data["Rows"]["Row"]:
                all_values.extend(extract_values(row))

            # convert into dataframe
            df = pd.DataFrame(all_values, columns=headers)
            df.insert(0, "Account", df.pop("Account"))
            df["Amount"] = pd.to_numeric(df["Amount"].str.replace(",", ""), errors="coerce")
            df["Balance"] = pd.to_numeric(df["Balance"].str.replace(",", ""), errors="coerce")

            # print(df)

            st.write("Clean Table:")
            st.dataframe(df)

            excel_binary = convert_df_to_excel(df)

            st.download_button(
                label="Download",
                data=excel_binary,
                file_name="qb_GL.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except:
            st.write("Raw Response:", response.text if response.text else "(empty response)")
