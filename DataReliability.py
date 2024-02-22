import os
import shutil
from zipfile import ZipFile
import pandas as pd

def find_fake_files(dataset_path, lowest_files=10, skip=False):
    # Load the combined dataset
    combined_dataset = pd.read_excel(dataset_path)

    # Gather basis for the most common labels
    label_columns = ['Consent_C', 'Consent_D', 'Consent_E']

    # Convert labels to lowercase and remove leading/trailing whitespaces, excluding NaN and empty values
    all_labels = combined_dataset[label_columns].applymap(lambda x: str(x).lower().strip() if pd.notna(x) and str(x).strip() else '').values.flatten()
    label_counts = pd.Series(all_labels).value_counts()

    # Remove the empty string from label counts
    label_counts = label_counts[label_counts.index != '']
    # Filter labels with counts less than 50
    label_counts = label_counts[label_counts >= 50]

    print("Top 3 consent violation labels:")
    print(label_counts.head(3))

    if skip:
        list1 = {}
        list2 = {}
        return list1, list2

    # Assign rankings based on label frequencies
    rankings = {}
    top_labels = label_counts.index[:3]
    for label in top_labels:
        rankings[label] = 1

    upper_labels = label_counts.index[3:5]
    for label in upper_labels:
        rankings[label] = 0.5

    lower_labels = label_counts.index[5:7]
    for label in lower_labels:
        rankings[label] = 0.25

    bottom_labels = label_counts.index[7:]
    for label in bottom_labels:
        rankings[label] = 0

    # Path to the zip file containing Labeled Reviews Excel files
    zip_path = r'C:\Users\lance\PycharmProjects\csc455_p2b\P2b\P2b\P2a Labels.zip'

    # Extract the contents of the zip file to a temporary folder
    temp_folder = 'temp_extracted'
    with ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_folder)

    # Path to the extracted Labeled Reviews Excel files
    excel_folder = os.path.join(temp_folder, 'P2a Labels')

    # Initialize variables for storing file scores
    file_scores = {}

    # Iterate through all 64 .xlsx files
    for i in range(1, 65):
        file_name = f'{i}.xlsx'
        file_path = os.path.join(excel_folder, file_name)

        # Check if the file exists
        if not os.path.exists(file_path):
            print(f"Skipping {file_name} - File not found.")
            continue

        # Read the current file
        try:
            current_file = pd.read_excel(file_path)
        except pd.errors.EmptyDataError:
            print(f"Skipping {file_name} - Empty DataFrame.")
            continue

        # Check if the file has exactly 5 columns
        if len(current_file.columns) != 5:
            print(f"Skipping {file_name} - Expected 5 columns, found {len(current_file.columns)} columns.")
            continue

        # Initialize variables for each file
        ranked_labels = 0
        label_count = 0

        # Iterate through each label column
        for label_column in current_file.columns:
            # Convert labels to lowercase and remove leading/trailing whitespaces
            current_labels = current_file[label_column].apply(lambda x: str(x).lower().strip() if pd.notna(x) and str(x).strip() else '')

            # Update variables based on the rankings
            ranked_labels += sum(rankings[label] for label in current_labels if label in rankings)
            label_count += len([label for label in current_labels if label in rankings])

        # Calculate the file score
        file_score = ranked_labels / label_count if label_count > 0 else 0

        # Store the file score
        file_scores[file_name] = file_score

    # Get the specified number of files with the lowest scores
    lowest_files = sorted(file_scores, key=file_scores.get)[:lowest_files]

    # Get the files that are not among the lowest
    non_lowest_files = [file for file in file_scores.keys() if file not in lowest_files]

    shutil.rmtree(temp_folder)

    return lowest_files, non_lowest_files
