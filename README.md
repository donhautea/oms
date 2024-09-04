# oms
Order Management System Tool

This tool allows you to process buying orders from Excel files and download a processed version with validated data and an updated filename based on predefined template of IMS file upload with the current Philippine time for filenaming convention.

# User Guide:

# Download Template:

Click the "Download Excel Template" button to get a ready-to-use template for entering your buying order details.
Upload File:

Upload your completed Excel file using the "Upload File" button. The file should have columns like "Order Type", "Stock", "Fund", etc.
Process and Download:

After uploading, the script will validate your data and process it.
Click "Download Processed File" to download the updated file with a dynamic filename (based on the current date and time in the Philippines).

# This tool ensures your data is consistent and easy to manage with automatic timestamping.


This Streamlit-based Python script processes buying orders from Excel files and allows users to download a processed version of their data. It features dynamic Excel file handling and provides templates and processed results with a filename that reflects the current date and time in Philippine time (UTC+8).

# Key Functionalities:
Download Excel Template: The app provides an option to download a predefined Excel template with columns like "Order Type", "Stock", "Fund", etc., for users to fill in their buying order details.

# File Upload & Processing: 
Users can upload an Excel file (either .xlsx or .xls) with buying order details. The script validates that the file contains the required columns and then processes the data by:

Converting broker names to uppercase.
Mapping broker names to their short forms using a reference CSV file (Broker_Master.csv).
Trimming and converting scheme names and stock tickers to uppercase.
Verifying scheme names against a second reference file (Scheme_Master.csv).
Filtering out invalid scheme names and creating a new processed DataFrame.
Dynamic Filename Generation: When users download the processed file, the filename is dynamically generated based on the current date and time in the format Buy_yyyymmddhhmm.xlsx, ensuring a unique filename each time.

# Philippine Timezone: 
The script ensures that the timestamp used for filenames is based on Philippine time (UTC+8) using the pytz library.

# Download Processed File: 
After processing the uploaded data, the app provides a download button to save the processed file as an Excel sheet.

# This tool is ideal for handling and processing buying orders in Excel format, ensuring consistency, validation, and ease of use with Philippine timezone-based timestamping.
