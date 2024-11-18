import pandas as pd
from collections import Counter


# Read existing warnings from Excel
def read_warnings_from_excel(file_path):
   df = pd.read_excel(file_path)
   warnings = df['Warning'].tolist()
   return warnings


# Count occurrences of each warning condition
def count_warning_occurrences(warnings):
   # Split combined warnings into individual warnings
   split_warnings = [warning.split(', ') for warning in warnings]
   # Flatten the list of warnings
   flat_warnings = [item for sublist in split_warnings for item in sublist]
   # Count occurrences of each warning
   warning_counts = dict(Counter(flat_warnings))
   return warning_counts


# Save warning counts to Excel
def save_warning_counts_to_excel(warning_counts, file_path):
   df = pd.DataFrame(warning_counts.items(), columns=['Warning', 'Occurrences'])
   df.to_excel(file_path, index=False)
   print(f"Warning counts saved to {file_path}")


# Example usage
warnings = read_warnings_from_excel('warnings_week_ex.xlsx')
warning_counts = count_warning_occurrences(warnings)
save_warning_counts_to_excel(warning_counts, 'tabelle_warning_counts.xlsx')