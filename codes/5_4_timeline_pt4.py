import pandas as pd


# Read reviews from Excel
def read_reviews_from_excel(file_path, sheet_name):
   df = pd.read_excel(file_path, sheet_name=sheet_name)
   reviews = []
   for _, row in df.iterrows():
       sentiment_quads = []
       for i in range(1, 9):  # Assuming there can be up to 8 sentiment quads per review
           aspect_term = row[f'aspect_term_{i}']
           aspect_category = row[f'aspect_category_{i}']
           sentiment_polarity = row[f'sentiment_polarity_{i}']
           opinion_term = row[f'opinion_term_{i}']
           if pd.notna(aspect_category) and pd.notna(sentiment_polarity) and pd.notna(opinion_term):
               sentiment_quads.append({'aspect_term': aspect_term, 'aspect_category': aspect_category, 'sentiment_polarity': sentiment_polarity, 'opinion_term': opinion_term})
       reviews.append(sentiment_quads)
   return reviews


# Check for four or more consecutive negative reviews in one week
def has_consecutive_negatives(reviews):
   consecutive_negatives = 0  # Counter for consecutive negative reviews
   max_consecutive_negatives = 4  # Threshold for consecutive negative reviews


   for week_reviews in reviews:
       for quad in week_reviews:
           if quad['sentiment_polarity'].lower() == 'negative':
               consecutive_negatives += 1
               if consecutive_negatives >= max_consecutive_negatives:
                   return True
           else:
               consecutive_negatives = 0  # Reset counter if a positive review is encountered


   return False


# Check if more than 50% of the reviews in one week are negative
def more_than_50_percent_negative(reviews):
   total_reviews = sum(1 for week_reviews in reviews for quad in week_reviews)
   num_negative_reviews = sum(1 for week_reviews in reviews for quad in week_reviews if quad['sentiment_polarity'].lower() == 'negative')
   if total_reviews != 0 and num_negative_reviews / total_reviews > 0.5:
       return True
   return False


# Check if more than 50% of negative reviews in a week are in the same category
def more_than_50_percent_same_category(reviews):
   category_counts = {}
   for week_reviews in reviews:
       for quad in week_reviews:
           if quad['sentiment_polarity'].lower() == 'negative':
               aspect_category = quad['aspect_category']
               category_counts[aspect_category] = category_counts.get(aspect_category, 0) + 1


   if category_counts:
       max_count_category = max(category_counts, key=category_counts.get)
       max_count = category_counts[max_count_category]
       num_negative_reviews = sum(category_counts.values())
       if max_count / num_negative_reviews > 0.5:
           return max_count_category


   return None


# Generate warnings and save to Excel
def generate_warnings_and_save(file_path, existing_file_path):
   warnings_data = []


   for week in range(1, 53):
       week_reviews = read_reviews_from_excel('reviewsfin.xlsm', sheet_name=f'Week{week}')
      
       warnings = []


       if has_consecutive_negatives(week_reviews):
           warnings.append('Four or more consecutive negative reviews in this week.')


       if more_than_50_percent_negative(week_reviews):
           warnings.append('More than 50% of the reviews in this week are negative.')


       max_count_category = more_than_50_percent_same_category(week_reviews)
       if max_count_category:
           warnings.append(f'More than 50% of negative reviews in this week are in the {max_count_category} category.')


       num_warnings = len(warnings)
       if warnings:
           warnings_data.append({'Week': week, 'Number of Warnings': num_warnings, 'Warnings': ', '.join(warnings)})
       else:
           warnings_data.append({'Week': week, 'Number of Warnings': 0, 'Warnings': 'No warnings for this week.'})


   # Read existing data from Excel
   existing_df = pd.read_excel(existing_file_path)


   # Convert warnings data to DataFrame
   warnings_df = pd.DataFrame(warnings_data)


   # Merge existing data with warnings data on the 'Week' column
   merged_df = pd.merge(existing_df, warnings_df, on='Week', how='left')


   # Save the merged data back to Excel
   merged_df.to_excel(file_path, index=False)
   print(f"Warnings saved to {file_path}")


# Example usage
generate_warnings_and_save('tabelle_warning_neg.xlsx', 'tabelle_weekly_neg.xlsx')