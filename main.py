from DataReliability import find_fake_files
from DatasetCombinerWithDuplicates import process_labeled_reviews_with_duplicates
from DatasetCombinerWithoutDuplicates import process_labeled_reviews_without_duplicates
from MajorityVoting import find_and_combine_duplicates
from NLTKAnalysis import analyze_and_visualize_wordcloud

# Call the function to get the first output that will be used to remove the 5 candidate files
# Define your zip file path to the labels you want to extract in zip_path variable
zip_path = r'C:\Users\lance\PycharmProjects\csc455_p2b\P2b\P2b\P2a Labels.zip'
output_directory = r'C:\Users\lance\PycharmProjects\csc455_p2b\output'
output_filename = 'final_dataset.xlsx'
process_labeled_reviews_without_duplicates(zip_path, output_directory, output_filename)

# Call the function to get top 10 fake candidate files
# Define your zip file path to the dataset of combined labels to analyze in dataset_path
dataset_path = r'C:\Users\lance\PycharmProjects\csc455_p2b\output\final_dataset.xlsx'
top_10_candidate_files, result_files_10 = find_fake_files(dataset_path, 10)
print("Result Files:", result_files_10)
print("The top 10 file candidates for being fake are: ", top_10_candidate_files)

# Call the function to remove the top 5 fake candidates
top_5_candidate_Files, result_files_5 = find_fake_files(dataset_path, 5)
print("Result Files:", result_files_5)
print("The top 5 fake file candidates removed from the final resulting dataset are: ", top_5_candidate_Files)

# Manually create new file folder with the 5 fake candidates removed

# Now use the file path with the zip containing the files with the 5 fake files removed
zip_path = r'C:\Users\lance\PycharmProjects\csc455_p2b\5_fake_files_removed\zip.zip'
output_directory = r'C:\Users\lance\PycharmProjects\csc455_p2b\output'
output_filename = 'final_dataset_without_5_fake_files.xlsx'
zip_name = 'zip'
process_labeled_reviews_with_duplicates(zip_path, output_directory, zip_name, output_filename)

# Now do the majority voting on the outputed dataset with the five removed fake file candidates
input_dataset_path = r'C:\Users\lance\PycharmProjects\csc455_p2b\output\final_dataset_without_5_fake_files.xlsx'
output_dataset_path = r'C:\Users\lance\PycharmProjects\csc455_p2b\output\majority_voted_dataset.xlsx'
find_and_combine_duplicates(input_dataset_path, output_dataset_path)

# Call the find_fake_file function but just to discover the top 3 most prominent labels in the new dataset
find_fake_files(output_dataset_path, 0, True)

# Do the NLTK Word analysis for word cloud

# Specify file path for your final_dataset_without_5_fake_files.xlsx
file_path = r'C:\Users\lance\PycharmProjects\csc455_p2b\output\final_dataset_without_5_fake_files.xlsx'
categories = ['Consent_C', 'Consent_D', 'Consent_E']

# Call the function
analyze_and_visualize_wordcloud(file_path, categories)
