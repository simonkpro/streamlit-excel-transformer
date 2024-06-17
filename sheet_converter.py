import streamlit as st
import pandas as pd
import requests
import re
import tempfile

# Function to get current exchange rates
def get_exchange_rate():
    api_url = "https://api.exchangerate-api.com/v4/latest/EUR"
    response = requests.get(api_url)
    data = response.json()
    return data['rates']

# Get the exchange rates once at the start
exchange_rates = get_exchange_rate()

# Function to convert currency to EUR using exchange rates
def convert_currency_to_eur(amount_str):
    try:
        # Ensure amount_str is a string
        if not isinstance(amount_str, str):
            amount_str = str(amount_str)

        # Detect the currency using regex
        if '£' in amount_str:
            detected_currency = 'GBP'
        elif 'USD$' in amount_str or '$' in amount_str or 'USD' in amount_str:
            detected_currency = 'USD'
        elif '€' in amount_str:
            detected_currency = 'EUR'
        elif 'GBP' in amount_str:
            detected_currency = 'GBP'
        else:
            detected_currency = 'EUR'  # Default to EUR if no symbol or abbreviation found

        # Extract amount from the amount_str
        amount = float(re.sub(r'[^\d.]+', '', amount_str))

        if detected_currency == 'EUR':
            return amount

        # Get the exchange rate for the detected currency
        exchange_rate = exchange_rates.get(detected_currency)
        if exchange_rate:
            return amount / exchange_rate
        else:
            st.warning(f"Exchange rate not found for {detected_currency}")
            return amount  # Return the amount as-is if exchange rate not found
    except Exception as e:
        st.error(f"Error converting currency: {e}")
        return 0

# Function to extract date from status using regex
def extract_date_from_status(status):
    try:
        # Ensure status is a string
        if not isinstance(status, str):
            status = str(status)

        # Look for date patterns like "dd/mm/yyyy"
        match = re.search(r'(\d{2}/\d{2}/\d{4})', status)
        if match:
            date_str = match.group(0)
            date_obj = pd.to_datetime(date_str, dayfirst=True)
            return date_obj.strftime('%Y-%m-%d 05:00')
        else:
            return "Date missing"
    except Exception as e:
        st.error(f"Error extracting date: {e}")
        return "Date missing"

# Streamlit interface
st.title("Excel Transformation Tool")

uploaded_file = st.file_uploader("Choose an input Excel file", type="xlsx")

if uploaded_file is not None:
    try:
        # Read the uploaded file
        df_input = pd.read_excel(uploaded_file)
        st.success("Input file read successfully.")

        # Check if 'Bank account' is present in the input DataFrame
        if 'Bank account' not in df_input.columns:
            st.error("'Bank account' column not found in the input sheet.")
        else:
            # Create the output DataFrame with the required columns
            columns = ["Amount", "Close Date", "Company Name", "Deal Description", "Deal Name", "Deal Owner", "Forecast Amount", "Create Date", "Invoice Number", "Bank account"]
            df_output = pd.DataFrame(columns=columns)

            # Map the input columns to the output columns and transform the data
            df_output["Amount"] = df_input["Amount"].apply(convert_currency_to_eur)
            df_output["Close Date"] = df_input["Status"].apply(extract_date_from_status)
            df_output["Company Name"] = df_input["Client"]
            df_output["Deal Description"] = df_input["Service"]
            df_output["Deal Name"] = df_input.apply(lambda row: f"{pd.to_datetime(row['Date']).strftime('%b-%y')} {row['Client']}", axis=1)
            df_output["Deal Owner"] = df_input["RESPONSABLE GESTION"]
            df_output["Forecast Amount"] = df_input["Amount"].apply(convert_currency_to_eur)
            df_output["Create Date"] = pd.to_datetime(df_input["Date"]).dt.strftime('%Y-%m-%d %H:%M')
            df_output["Invoice Number"] = df_input["Invoice number"]
            df_output["Bank account"] = df_input["Bank account"]

            # Create a temporary file to save the output
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                df_output.to_excel(tmp.name, index=False)
                tmp.seek(0)
                st.success("Transformation complete.")
                st.download_button(
                    label="Download Transformed Excel",
                    data=tmp.read(),
                    file_name="output_sheet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"Error processing the file: {e}")
