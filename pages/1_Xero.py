import streamlit as st
import requests
import pandas as pd
from io import BytesIO
from datetime import datetime, timezone
import time

def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Sheet1")
            return output.getvalue()            

def formatDate(date):
    timestamp_ms = int(date.split("(")[1].split("+")[0])
    clean_date = datetime.fromtimestamp(timestamp_ms / 1000, tz=timezone.utc)
    return clean_date.date().isoformat()

all_journals = []
headers = ["Date", "Journal Number", "Account Code", "Account Type", "Account Name", "Reference", "Description", "Net Amount", "Tax Amount", "Tax Type", "Tax Name", "Gross Amount"]

st.title("Xero GL Report")

if st.button("Run workflow"):
    with st.spinner("Running workflow..."):
        webhook = "https://hook.us2.make.com/yv6elkcn4q7f6yf9zhzruscwy4x2eau1"  # your webhook
        JournalNumber = 0
        while True:
            # Send payload to Make
            response = requests.post(webhook, json={
                "value": "xero",
                "offset": str(JournalNumber)})
            
            if response.status_code != 200:
                st.write("Error:", response.text)
                break

            try:
                data = response.json()
                journals = data["Journals"]

                if len(journals) == 0:  # stop when no more journals
                    print("no more journals")
                    break

                else:
                    for journal in journals:
                        Date = formatDate(journal["JournalDate"])
                        JournalNumber = journal["JournalNumber"]
                        Reference = journal.get("Reference")

                        # Loop each journal line
                        for line in journal["JournalLines"]:
                            AccountCode = line["AccountCode"]
                            AccountType = line["AccountType"]
                            AccountName = line["AccountName"]
                            Description = line.get("Description")
                            NetAmount = line["NetAmount"]
                            TaxAmount = line.get("TaxAmount")
                            TaxType = line.get("TaxType")
                            TaxName = line.get("TaxName")
                            GrossAmount = line["GrossAmount"] #NetAmount + TaxAmount

                            all_journals.append([Date, JournalNumber, AccountCode, AccountType, AccountName, Reference, Description, NetAmount, TaxAmount, TaxType, TaxName, GrossAmount])

            except:
                 print("error")           
                 print(data)
       
            time.sleep(2)  # wait 1 second between requests

        df = pd.DataFrame(all_journals, columns=headers)
        df = df.sort_values(by=["Account Code","Date"])
        st.write("Clean Table:")
        st.dataframe(df)

        excel_binary = convert_df_to_excel(df)

        st.download_button(
            label="Download",
            data=excel_binary,
            file_name="xero_journal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )