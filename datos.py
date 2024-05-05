import pandas as pd
from openpyxl import load_workbook
from collections import Counter
import nltk
from nltk.stem import SnowballStemmer
from sklearn.feature_extraction.text import TfidfVectorizer


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
words = [word.replace("!", "").replace("?", "") for word in words] # eliminar signos de exclamaci'on
words_to_remove = ['de', 'la','el','que','lo','en','me','por','nos','y','a','ya','con','para','mi','es','un','los','una','las','se','te']

# Remove specific words from the list
words = [word for word in words if word not in words_to_remove]

print(len(words))
# Now you can work with your DataFrame
word_counts = Counter(words)

# Output the word counts
print(word_counts)
