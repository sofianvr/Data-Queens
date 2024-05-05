import re
import pandas as pd
from openpyxl import load_workbook
from collections import Counter
import nltk
from nltk.stem import SnowballStemmer
from sklearn.feature_extraction.text import TfidfVectorizer
import csv


# Function to read Excel file using openpyxl and return a DataFrame
def read_excel_with_openpyxl(file_path):
    # Load the Excel file using openpyxl
    wb = load_workbook(filename=file_path)
    # Get the active sheet
    sheet = wb.active
    # Get the data from the sheet
    data = sheet.values
    # Convert data to DataFrame
    df = pd.DataFrame(data)
    return df

# Read the Excel file into a DataFrame using the custom function
df = read_excel_with_openpyxl(r'C:\Users\amvc6\Documents\Amanda\Uni\Datathon 2024\Datathon2024_heybanco.xlsm')
comments = df.iloc[:,2]
#print(comments)
words = comments.str.split(r'\s*,\s*|\s+')
words = [item for sublist in words for item in sublist]
print(len(words))
word_counts = {}

words = [word.lower() for word in words] # convertir las palabras a minúsculas
words = [word.replace("!", "").replace("?", "") for word in words] # eliminar signos de exclamacion
words_to_remove = ['de', 'la','el','que','lo','en','me','por','nos','y','a','ya','con','para','mi','es','un','los','una','las','se','te']

# Remove specific words from the list
words = [word for word in words if word not in words_to_remove]

print(len(words))
words = ' '.join(words)
print(type(words))
# Now you can work with your DataFrame
#word_counts = Counter(words)

# Output the word counts
#print(word_counts)


emoji_pattern = r'[\U0001F300-\U0001F5FF\U0001F600-\U0001F64F\U0001F680-\U0001F6FF\U0001F700-\U0001F77F\U0001F780-\U0001F7FF\U0001F800-\U0001F8FF\U0001F900-\U0001F9FF\U0001FA00-\U0001FA6F\U0001FA70-\U0001FAFF\U00002702-\U000027B0\U000024C2-\U0001F251\U0001F004\U0001F0CF\U0001F170-\U0001F251\U0001F1E6-\U0001F1FF]+'

# Find all emoji characters using the pattern
emojis = re.findall(emoji_pattern, words)

# Count the occurrences of each emoji
emoji_counts = Counter(emojis)
print(type(emoji_counts))
# Output the emoji counts
#print("Cuenta de emojis", emoji_counts)

csv_filename = "emojis_cuenta.csv"

with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
    csv_writer = csv.writer(csvfile)
    csv_writer.writerow(['Emoji', 'Count'])  # Write header
    for word, count in emoji_counts.items():
        csv_writer.writerow([word, count])

print(f"Counter data has been exported to '{csv_filename}'.")