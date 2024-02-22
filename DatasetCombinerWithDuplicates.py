import os
import pandas as pd
from zipfile import ZipFile
import shutil

def process_labeled_reviews_with_duplicates(zip_path, output_directory, zip_name, output_filename='final_dataset.xlsx'):
    # Extract the contents of the zip file to a temporary folder
    temp_folder = 'temp_extracted'
    with ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_folder)

    # Path to the extracted Labeled Reviews Excel files
    excel_folder = os.path.join(temp_folder, zip_name)

    # Initialize an empty dataframe to store the combined data
    combined_dataset = pd.DataFrame(columns=['Header', 'Body', 'Consent_C', 'Consent_D', 'Consent_E'])

    # Loop through each file in the folder, sorting by numerical value in the filename
    for file_name in sorted(os.listdir(excel_folder), key=lambda x: int(os.path.splitext(x)[0])):
        file_path = os.path.join(excel_folder, file_name)
        if file_name.endswith('.xlsx'):
            # Read the Excel file into a pandas dataframe, skipping the first row
            df = pd.read_excel(file_path, header=None, skiprows=1)

            # Check if the dataframe has the expected number of columns
            # if they don't have 5 columns then skip the Excel file
            if df.shape[1] != 5:
                print(f"Skipping file {file_name} due to unexpected column count.")
                continue

            # Process the dataframe to ensure consistent formatting
            num_columns = df.shape[1]
            df.columns = ['Header', 'Body'] + [f'Consent_{chr(ord("C") + i)}' for i in range(num_columns - 2)]

            # Add rows with at least one label in C, D, or E columns (excluding rows with all empty labels)
            mask = (df['Consent_C'].notna()) | (df['Consent_D'].notna()) | (df['Consent_E'].notna())
            df_filtered = df[mask]

            # Add all rows to the combined dataset, including duplicates
            combined_dataset = pd.concat([combined_dataset, df_filtered], ignore_index=True)

    # Replace blanks with "no violation"
    combined_dataset[['Consent_C', 'Consent_D', 'Consent_E']] = combined_dataset[['Consent_C', 'Consent_D', 'Consent_E']].replace('', 'no violation')

    # Create the output directory if it doesn't exist
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Save the combined dataset to a new Excel file in the specified directory
    combined_dataset.to_excel(os.path.join(output_directory, output_filename), index=False)

    shutil.rmtree(temp_folder)
