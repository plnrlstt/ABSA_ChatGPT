import pandas as pd
import openai

# Set your OpenAI API key
openai.api_key = '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'

def generate_weekly_summary(reviews):
    negative_reviews = []
    
    # Loop through the reviews to identify negative ones
    for review in reviews:
        for quad in review['sentiment_quads']:
            if quad['sentiment_polarity'] == 'negative':
                # Check if aspect term is not NULL
                if quad['aspect_term'] is not None:
                    negative_reviews.append((quad['aspect_category'], quad['aspect_term'], quad['opinion_term']))

    # Generate the weekly summary
    if len(negative_reviews) == 0:
        return "No negative reviews for this week."
    elif len(negative_reviews) == 1:
        aspect_category, aspect_term, opinion_term = negative_reviews[0]
        if aspect_term:
            return f"In this week, there was one negative review, and that was in the aspect category of {aspect_category}. The {aspect_term} was described as {opinion_term}."
        else:
            return f"In this week, there was one negative review, and that was in the aspect category of {aspect_category}. The review mentioned: {opinion_term}."
    else:
        summary = f"In this week, there were {len(negative_reviews)} negative reviews. The negative aspects included:\n"
        for aspect_category, aspect_term, opinion_term in negative_reviews:
            if aspect_term:
                summary += f"- {aspect_category}: {aspect_term} was described as {opinion_term}.\n"
            else:
                summary += f"- {aspect_category}: {opinion_term}.\n"
        return summary
        
# Save weekly summaries to Excel
def save_weekly_summaries_to_excel(file_path):
    summaries = generate_weekly_summaries()
    
    # Create lists to store data for each column
    week_list = []
    summary_list = []
    aspect_description_list = []
    
    # Loop through the summaries and extract data for each column
    for week, summary in enumerate(summaries, start=1):
        lines = summary.split('\n')
        week_list.append(week)
        summary_list.append(lines[0])
        aspect_description = ''
        for line in lines[1:]:
            aspect_description += line + '\n'
        aspect_description_list.append(aspect_description.strip())
    
    # Create a DataFrame from the extracted data
    df = pd.DataFrame({
        'Week': week_list,
        'Summary': summary_list,
        'Aspect and Description': aspect_description_list
    })
    
    # Save the DataFrame to Excel
    df.to_excel(file_path, index=False)
    print(f"Weekly summaries saved to {file_path}")

# Example usage
save_weekly_summaries_to_excel('sum_week_ex.xlsx')
