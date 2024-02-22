import pandas as pd
from openpyxl import load_workbook

def majority_vote(labels):
    # Perform majority voting
    # If there is a tie or all labels are empty, return an empty string
    label_counts = pd.Series(labels).value_counts()

    if len(label_counts) == 0:
        return ''
    elif len(label_counts) == 1:
        return label_counts.idxmax()
    elif label_counts.iloc[0] == label_counts.iloc[1]:
        return ''
    else:
        return label_counts.idxmax()

def combine_labels(labels_list):
    # Combine labels using majority voting
    return [majority_vote(labels) for labels in zip(*labels_list)]

def find_and_combine_duplicates(input_path, output_path):
    # Load the input dataset
    input_dataset = pd.read_excel(input_path)

    # Initialize a dictionary to store labels for each review
    review_labels = {}

    # Iterate through each row in the input dataset
    for index, row in input_dataset.iterrows():
        header = row['Header']
        labels = row[['Consent_C', 'Consent_D', 'Consent_E']].tolist()

        # If the review is not in the dictionary, add it
        if header not in review_labels:
            review_labels[header] = []

        # Add the labels for the current row to the dictionary
        review_labels[header].append(labels)

    # Initialize a new dataframe for the final dataset
    final_dataset = pd.DataFrame(columns=pd.Index(['Header', 'Body', 'Consent_C', 'Consent_D', 'Consent_E']))

    # Iterate through each review and its labels
    for header, labels_list in review_labels.items():
        # Combine labels for each review using majority voting
        combined_labels = combine_labels(labels_list)

        # Append the combined labels to the final dataset
        final_dataset = pd.concat([final_dataset, pd.DataFrame({'Header': [header], 'Body': [''],
                                                                'Consent_C': [combined_labels[0]],
                                                                'Consent_D': [combined_labels[1]],
                                                                'Consent_E': [combined_labels[2]]})],
                                  ignore_index=True)

    # Save the final dataset to a new Excel file using openpyxl
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        final_dataset.to_excel(writer, index=False, sheet_name='Sheet1')
