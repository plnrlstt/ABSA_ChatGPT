import pandas as pd
import openai


# Set your OpenAI API key
openai.api_key = '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'


def generate_weekly_summary(reviews):
   negative_reviews_count = 0
  
   # Loop through the reviews to identify negative ones
   for review in reviews:
       for quad in review['sentiment_quads']:
           if quad['sentiment_polarity'] == 'negative':
               negative_reviews_count += 1


   return negative_reviews_count


# Read reviews from Excel
def read_reviews_from_excel(file_path, sheet_name):
   df = pd.read_excel(file_path, sheet_name=sheet_name)
   reviews = []
   for _, row in df.iterrows():
       sentiment_quads = []
       for i in range(1, 9):  # Assuming there can be up to 8 sentiment quads per review
           aspect_category = row[f'aspect_category_{i}']
           sentiment_polarity = row[f'sentiment_polarity_{i}']
           if pd.notna(aspect_category) and pd.notna(sentiment_polarity):
               sentiment_quads.append({'aspect_category': aspect_category, 'sentiment_polarity': sentiment_polarity})
       reviews.append({'sentiment_quads': sentiment_quads})
   return reviews


# Save weekly summaries to Excel
def save_weekly_summaries_to_excel(file_path):
   week_summaries = []
   for week in range(1, 53):
       week_reviews = read_reviews_from_excel('reviewsfin.xlsm', sheet_name=f'Week{week}')
       negative_reviews_count = generate_weekly_summary(week_reviews)
       week_summaries.append({'Week': week, 'Number of Negative Reviews': negative_reviews_count})


   df = pd.DataFrame(week_summaries)
   df.to_excel(file_path, index=False)
   print(f"Weekly summaries saved to {file_path}")


# Example usage
save_weekly_summaries_to_excel('tabelle_weekly_neg.xlsx')
