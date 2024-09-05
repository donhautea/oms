import streamlit as st
import pandas as pd
import os

# Function to load Excel file and return as DataFrame
def load_data(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

# Main function to run the app
def main():
    st.title("AT Equity Data Viewer")

    # Sidebar section
    st.sidebar.header("File Selection")
    
    # Specify the folder path where the file is located
    folder_path = r"D:\git_repository\Trader_upload\Source_Files"
    
    # List all files in the folder
    files = [f for f in os.listdir(folder_path) if f.endswith('.xls') or f.endswith('.xlsx')]
    
    # Allow the user to select a file
    selected_file = st.sidebar.selectbox("Select an Excel file", files)
    
    # Load the selected file
    if selected_file:
        file_path = os.path.join(folder_path, selected_file)
        df = load_data(file_path)
        
        if df is not None:
            st.write(f"Displaying data from {selected_file}")
            st.dataframe(df)

if __name__ == "__main__":
    main()
