import os
from dotenv import load_dotenv
import pandas as pd
from openai import OpenAI, OpenAIError

# Initialize OpenAI client - you need to create a file called .env in the same directory as this script. The .env file should contain the following line: OPENAI_API_KEY="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"
load_dotenv()

client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
)

# Load questions from Excel sheet. Change to the name of the file you want to use
excel_file_path = "reviewsfin.xlsm"

# Fetch sheet names dynamically from the Excel file
sheet_names = pd.ExcelFile(excel_file_path).sheet_names

# Initialize an empty dictionary to hold the DataFrames
dfs = {}

# Read each sheet and store it in the dictionary
for sheet_name in sheet_names:
    dfs[sheet_name] = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# Initialize an empty dictionary to hold the results for each week
results = {}

# Loop through each sheet
for sheet_name, df in dfs.items():
    results[sheet_name] = []  # Initialize list to hold results for this week
    try:
        total_reviews = len(df)  # Total number of reviews in the sheet

        print(f"### {sheet_name} ###")  # Print header for each week
        
        # Our ChatGPT prompt
        for index, row in df.iterrows():
            question = """Identify the aspect term, the sentiment polarity, and the opinion term for the following sentences. Assign them to the appropriate category out of [location general, food prices, food quality, ambience general, service general, restaurant prices, drinks prices, restaurant miscellaneous, drinks quality, drinks style_options, restaurant general, food style_options].
                    Here are examples of how the results are supposed to look like for all the different categories:
                    
                    The location is perfect .
                    Aspect term: location
                    Aspect category: location general
                    Sentiment polarity: positive
                    Opinion term: perfect

                    But the pizza is way to expensive .
                    Aspect term: pizza
                    Aspect category: food prices
                    Sentiment polarity: negative
                    Opinion term: expensive

                    The food was good .
                    Aspect term: food
                    Aspect category: food quality
                    Sentiment polarity: positive
                    Opinion term: good

                    The owner and staff are all Japanese as well and that adds to the entire ambiance .
                    Aspect term: ambience
                    Aspect category: ambience general
                    Sentiment polarity: positive
                    Opinion term: adds

                    The owner truly caters to all your needs .
                    Aspect term: owner
                    Aspect category: service general
                    Sentiment polarity: positive
                    Opinion term: caters to all your needs

                    This place is a great bargain .
                    Aspect term: place
                    Aspect category: restaurant prices
                    Sentiment polarity: positive
                    Opinion term: great bargain

                    To start off , approximately 8-10 oz of orange juice will cost you $ 3 .
                    Aspect term: orange juice
                    Aspect category: drinks prices
                    Sentiment polarity: negative
                    Opinion term: cost you $ 3

                    A great place to meet up for some food and drinks ...
                    Aspect term: place
                    Aspect category: restaurant miscellaneous
                    Sentiment polarity: positive
                    Opinion term: great

                    The have a great cocktail with Citrus Vodka and lemon and lime juice and mint leaves that is to die for !
                    Aspect term: cocktail with Citrus Vodka and lemon and lime juice and mint leaves
                    Aspect category: drinks quality
                    Sentiment polarity: positive
                    Opinion term: great

                    Very good wine choices .
                    Aspect term: wine choices
                    Aspect category: drinks style_options
                    Sentiment polarity: positive
                    Opinion term: good

                    This place doesn 't make any sense .
                    Aspect term: place
                    Aspect category: restaurant general
                    Sentiment polarity: negative
                    Opinion term: doesn 't make any sense

                    With the great variety on the menu , I eat here often and never get bored .
                    Aspect term: menu
                    Aspect category: food style_options
                    Sentiment polarity: positive
                    Opinion term: great variety

                    If the aspect term is not specifically stated, it should be "NULL" like this:
                    
                    We asked for beverages and never received them .
                    Aspect term: NULL
                    Aspect category: service general
                    Sentiment polarity: negative
                    Opinion term: never received
                               
                    Lastly, if the sentence has multiple aspect terms, analyze them separately. The result for 2 aspect terms should look like this:

                    Decor is nice though service can be spotty .
                    Aspect term 1: Decor
                    Aspect category 1: ambience general
                    Sentiment polarity 1: positive
                    Opinion term 1: nice

                    Aspect term 2: service
                    Aspect category 2: service general
                    Sentiment polarity 2: negative
                    Opinion term 2: spotty

                    
                    Do the same for the following sentence: """  + row['Review']

            # Generate chat completion
            chat_completion = client.chat.completions.create(
                messages=[
                    {
                        "role": "user",
                        "content": question,
                    }
                ],
                model="gpt-3.5-turbo",
            )

            # Output from chatgpt
            if hasattr(chat_completion, 'choices') and chat_completion.choices:
                message_content = chat_completion.choices[0].message.content
                
                # Splitting the message content to extract individual parts
                parts = message_content.split('\n')

                # Initialize variables to store extracted information
                aspect_terms = []
                aspect_categories = []
                sentiment_polarities = []
                opinion_terms = []

                # Loop through each line of the message content
                for part in parts:
                    if part.startswith('Aspect term'):
                        aspect_terms.append(part.split(': ')[1])
                    elif part.startswith('Aspect category'):
                        aspect_categories.append(part.split(': ')[1])
                    elif part.startswith('Sentiment polarity'):
                        sentiment_polarities.append(part.split(': ')[1])
                    elif part.startswith('Opinion term'):
                        opinion_terms.append(part.split(': ')[1])

                # Append the extracted information to the list for this week
                for i in range(max(len(aspect_terms), len(aspect_categories), len(sentiment_polarities), len(opinion_terms))):
                    aspect_term = aspect_terms[i] if i < len(aspect_terms) else ''
                    aspect_category = aspect_categories[i] if i < len(aspect_categories) else ''
                    sentiment_polarity = sentiment_polarities[i] if i < len(sentiment_polarities) else ''
                    opinion_term = opinion_terms[i] if i < len(opinion_terms) else ''
                    results[sheet_name].append((row['Review'], aspect_term, aspect_category, sentiment_polarity, opinion_term))

            else:
                print("Error: No response content received from OpenAI.")

        print(f"Sheet '{sheet_name}' completed.")  # Print statement after completing the sheet

    except OpenAIError as e:
        print("OpenAI Error:", e)

# Save the results to a new sheet in the output Excel file
output_file_path = "output_all_weeks.xlsx"
with pd.ExcelWriter(output_file_path) as writer:
    for sheet_name, result_list in results.items():
        df = pd.DataFrame(result_list, columns=['Review', 'Aspect Term', 'Aspect Category', 'Sentiment Polarity', 'Opinion Term'])
        df.to_excel(writer, sheet_name=sheet_name, index=False)  # index=False to exclude the row numbers

print()
print(f"Results for all weeks saved to {output_file_path}")
print()
