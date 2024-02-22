import os
import pandas as pd
from zipfile import ZipFile
import shutil

def process_labeled_reviews_without_duplicates(zip_path, output_directory, output_filename='final_dataset.xlsx'):
    # Extract the contents of the zip file to a temporary folder
    temp_folder = 'temp_extracted'
    with ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_folder)

    # Path to the extracted Labeled Reviews Excel files
    excel_folder = os.path.join(temp_folder, 'P2a Labels')

    # Initialize an empty dataframe to store the combined data
    combined_dataset = pd.DataFrame(columns=['Header', 'Body', 'Consent_C', 'Consent_D', 'Consent_E'])

    # Set to keep track of headers already added to the combined dataset
    added_headers = set()

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

            # Check for duplicates based on the 'Header' column
            duplicates = df_filtered[df_filtered.duplicated('Header')]

            # Only add rows that are not duplicates and not already in the added_headers set
            df_unique = df_filtered[~df_filtered['Header'].isin(added_headers)]

            # Add unique rows to the combined dataset
            combined_dataset = pd.concat([combined_dataset, df_unique], ignore_index=True)

            # Update the set of added headers
            added_headers.update(df_unique['Header'].tolist())

    # Replace blanks with "no violation"
    combined_dataset[['Consent_C', 'Consent_D', 'Consent_E']] = combined_dataset[['Consent_C', 'Consent_D', 'Consent_E']].replace('', 'no violation')

    # Extract the first 2500 rows
    # final_dataset = combined_dataset.head(2500)
    final_dataset = combined_dataset

    # Create the output directory if it doesn't exist
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Save the combined dataset to a new Excel file in the specified directory
    final_dataset.to_excel(os.path.join(output_directory, output_filename), index=False)

    shutil.rmtree(temp_folder)
