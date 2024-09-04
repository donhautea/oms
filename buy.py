import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Function to create and provide a download link for the Excel template
def generate_excel_template():
    # Define the columns for the template
    template_columns = ['Order Type', 'Stock', 'Fund', 'Shares', 'Price Limit', 'Value', 'Classification', 'Broker', 'Remarks']

    # Create an empty DataFrame with these columns
    template_df = pd.DataFrame(columns=template_columns)

    # Write the DataFrame to a BytesIO object (acts like a file)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        template_df.to_excel(writer, index=False, sheet_name='Template')

    # Set the pointer of the BytesIO object to the beginning
    output.seek(0)
    
    return output

# Function to create a downloadable Excel file
def generate_excel_file(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Processed Data')
    output.seek(0)
    return output

# Streamlit app
def main():
    # Set the title of the application
    st.title("Buying Order Excel Processor")

    # Generate current timestamp for the default filename
    current_time = datetime.now().strftime("%Y%m%d%H%M")
    default_output_filename = f"Buy_{current_time}"

    # Sidebar for downloading the Excel template
    st.sidebar.title("Download Template")
    template = generate_excel_template()
    
    st.sidebar.download_button(
        label="Download Excel Template",
        data=template,
        file_name="excel_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Read the Broker_Master.csv from the local directory
    broker_master_path = "Broker_Master.csv"  # Adjust the path if necessary
    scheme_master_path = "Scheme_Master.csv"  # Adjust the path if necessary
    
    try:
        broker_master_df = pd.read_csv(broker_master_path)
    except FileNotFoundError:
        st.error(f"Broker_Master.csv not found at {broker_master_path}")
        return

    try:
        scheme_master_df = pd.read_csv(scheme_master_path)
        # Drop all columns except 'Scheme Short Name' and 'Scheme Name'
        scheme_master_df = scheme_master_df[['Scheme Short Name', 'Scheme Name']]
    except FileNotFoundError:
        st.error(f"Scheme_Master.csv not found at {scheme_master_path}")
        return

    # Sidebar for uploading the file
    st.sidebar.title("Upload File")
    uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

    # Sidebar for output file name input with default filename
    output_filename = st.sidebar.text_input("Enter output filename (without extension)", value=default_output_filename)

    if uploaded_file:
        # Load the uploaded Excel file into a DataFrame
        df = pd.read_excel(uploaded_file)

        # Show a preview of the uploaded data
        st.write("Uploaded data preview:")
        st.dataframe(df.head())

        # Ensure that the required columns are present
        required_columns = ['Order Type', 'Stock', 'Fund', 'Shares', 'Price Limit', 'Value', 'Classification', 'Broker', 'Remarks']
        if all(col in df.columns for col in required_columns):
            # Process the file
            new_df = process_file(df, broker_master_df, scheme_master_df)

            # Show the processed data
            st.write("Processed data preview:")
            st.dataframe(new_df.head())

            # Provide a download button for the processed file
            processed_file = generate_excel_file(new_df)
            st.download_button(
                label="Download Processed File",
                data=processed_file,
                file_name=f"{output_filename}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("The uploaded file is missing some required columns.")
    else:
        st.warning("Please upload an Excel file.")

# Function to process the uploaded file and create a new DataFrame
def process_file(df, broker_master_df, scheme_master_df):
    # Convert Broker to uppercase before merging
    df['Broker'] = df['Broker'].str.upper()

    # Merge the df with the Broker_Master.csv to replace BrokerShortname with Shortname
    df = df.merge(broker_master_df[['Code', 'Shortname']], left_on='Broker', right_on='Code', how='left')

    # Replace Broker with Shortname where available
    df['BrokerShortname'] = df['Shortname'].fillna(df['Broker'])

    # Trim and convert Scheme Short Name to uppercase, and ensure it's in Scheme_Master.csv
    df['SchemeShortName'] = df['Fund'].str.strip().str.upper()

    # Convert Stock to uppercase for ISIN mapping
    df['Stock'] = df['Stock'].str.upper()

    # Check if Scheme Short Name exists in Scheme_Master.csv
    valid_schemes = scheme_master_df['Scheme Short Name'].str.upper().tolist()
    df['SchemeShortNameValid'] = df['SchemeShortName'].isin(valid_schemes)

    # Filter only valid schemes
    if not df['SchemeShortNameValid'].all():
        st.warning("Some Scheme Short Names are not found in Scheme_Master.csv. They will be excluded from the processed data.")

    df = df[df['SchemeShortNameValid']]

    # Creating the new DataFrame with mapped columns
    new_df = pd.DataFrame({
        'AssetClassification': ['Equity'] * len(df),
        'SchemeShortName': df['SchemeShortName'],
        'ISIN': df['Stock'],  # ISIN is now uppercase from Stock
        'InstrumentHoldingType': df['Classification'],
        'BrokerShortname': df['BrokerShortname'],
        'ExchangeShortName': ['PSE'] * len(df),
        'CounterParty': [''] * len(df),
        'TransactionType': df['Order Type'],
        'OrderQuantity': df['Shares'],
        'OrderPrice': df['Price Limit'],
        'YTM': [''] * len(df),
        'Validity': ['GFD'] * len(df),
        'Remarks': df['Remarks']
    })
    return new_df

if __name__ == "__main__":
    main()
