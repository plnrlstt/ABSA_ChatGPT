import pandas as pd
import openai

# Set your OpenAI API key
openai.api_key = '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'


def generate_monthly_summary(weekly_summaries):
   weeks_per_month = 4
   monthly_negative_reviews = []


   for i in range(0, len(weekly_summaries), weeks_per_month):
       month_reviews = weekly_summaries[i:i+weeks_per_month]
       monthly_negative_reviews.append(sum(month_reviews))


   return monthly_negative_reviews


# Read weekly summaries from Excel
def read_weekly_summaries_from_excel(file_path):
   df = pd.read_excel(file_path)
   return df['Number of Negative Reviews'].tolist()


# Save monthly summaries to Excel
def save_monthly_summaries_to_excel(file_path):
   weekly_summaries = read_weekly_summaries_from_excel('tabelle_weekly_neg.xlsx')
   monthly_negative_reviews = generate_monthly_summary(weekly_summaries)


   month_summaries = [{'Month': f'{i+1}', 'Number of Negative Reviews': monthly_negative_reviews[i]} for i in range(len(monthly_negative_reviews))]


   df = pd.DataFrame(month_summaries)
   df.to_excel(file_path, index=False)
   print(f"Monthly summaries saved to {file_path}")


# Example usage
save_monthly_summaries_to_excel('tabelle_monthly_neg.xlsx')
