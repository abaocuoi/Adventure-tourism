# import pandas as pd
# from textblob import TextBlob

# # Load the Excel file
# file_path = 're-test.xlsx'
# df = pd.read_excel('re-test.xlsx', engine='openpyxl')

# # Assuming the comments are in the first column
# comments = df.iloc[:, 0]

# # Function to perform sentiment analysis
# def analyze_sentiment(comment):
#     if isinstance(comment, str):
#         analysis = TextBlob(comment)
#         return analysis.sentiment.polarity
#     else:
#         return None

# # Apply sentiment analysis to each comment
# df['Sentiment'] = comments.apply(analyze_sentiment)

# # Save the results to a new Excel file
# df.to_excel('sentiment_analysis_results.xlsx', index=False)


import pandas as pd
from textblob import TextBlob

# Load the Excel file
df = pd.read_excel('re-test.xlsx', engine='openpyxl')

# Assuming the comments are in the first column
comments = df.iloc[:, 1]

# Function to compute sentiment polarity
def get_sentiment(text):
    if isinstance(text, str):
        return TextBlob(text).sentiment.polarity
    return None

# Apply sentiment analysis
df['Sentiment'] = comments.apply(get_sentiment)

# Save the updated DataFrame to a new Excel file
df.to_excel('sentiment_scored.xlsx', index=False)
