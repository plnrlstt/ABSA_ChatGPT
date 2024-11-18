import pandas as pd

def merge_duplicates(df):
    # Initialize a dictionary to store information for each unique sentence
    unique_sentences = {}

    for index, row in df.iterrows():
        review = row['Review']
        if review not in unique_sentences:
            # Store information for the first occurrence of the sentence
            unique_sentences[review] = {
                'Review': review,
                '': '',
                'Aspect Term': row['Aspect Term'],
                'Aspect Category': row['Aspect Category'],
                'Sentiment Polarity': row['Sentiment Polarity'],
                'Opinion Term': row['Opinion Term'],
                '': '',
                'Aspect Term 2': '',
                'Aspect Category 2': '',
                'Sentiment Polarity 2': '',
                'Opinion Term 2': '',
                '': '',
                'Aspect Term 3': '',
                'Aspect Category 3': '',
                'Sentiment Polarity 3': '',
                'Opinion Term 3': '',
                '': '',
                'Aspect Term 4': '',
                'Aspect Category 4': '',
                'Sentiment Polarity 4': '',
                'Opinion Term 4': '',
                '': '',
                'Aspect Term 5': '',
                'Aspect Category 5': '',
                'Sentiment Polarity 5': '',
                'Opinion Term 5': '',
                '': '',
                'Aspect Term 6': '',
                'Aspect Category 6': '',
                'Sentiment Polarity 6': '',
                'Opinion Term 6': '',
                '': '',
                'Aspect Term 7': '',
                'Aspect Category 7': '',
                'Sentiment Polarity 7': '',
                'Opinion Term 7': '',
                '': '',
                'Aspect Term 8': '',
                'Aspect Category 8': '',
                'Sentiment Polarity 8': '',
                'Opinion Term 8': ''
            }
        else:
            # Find the first available slot for storing information
            for i in range(2, 9):
                if unique_sentences[review][f'Aspect Term {i}'] == '':
                    # Store information for the next available slot
                    unique_sentences[review][f'Aspect Term {i}'] = row['Aspect Term']
                    unique_sentences[review][f'Aspect Category {i}'] = row['Aspect Category']
                    unique_sentences[review][f'Sentiment Polarity {i}'] = row['Sentiment Polarity']
                    unique_sentences[review][f'Opinion Term {i}'] = row['Opinion Term']
                    unique_sentences[review][f''] = ''
                    break              

    # Create DataFrame from unique_sentences dictionary
    unique_df = pd.DataFrame.from_dict(unique_sentences, orient='index')

    return unique_df

# Read the Excel file
file_path = 'output_all_weeks.xlsx'
excel_data = pd.read_excel(file_path, sheet_name=None)

# Process each sheet in the Excel file
for sheet_name, df in excel_data.items():
    # Merge duplicates for each sheet
    excel_data[sheet_name] = merge_duplicates(df)

# Save the modified Excel file
with pd.ExcelWriter('output_sorted.xlsx') as writer:
    for sheet_name, df in excel_data.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
