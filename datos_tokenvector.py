import pandas as pd
from openpyxl import load_workbook
from collections import Counter
import nltk
from nltk.stem import SnowballStemmer
from sklearn.feature_extraction.text import TfidfVectorizer
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.collocations import TrigramCollocationFinder
from nltk.metrics import TrigramAssocMeasures
from nltk.util import ngrams


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

def delete_comments(df, word):
    # Filter rows based on whether the word is in any column
    filtered_df = df[~df.apply(lambda row: row.astype(str).str.contains(word).any(), axis=1)]
    return filtered_df


# Read the Excel file into a DataFrame using the custom function
df = read_excel_with_openpyxl(r'C:\Users\amvc6\Documents\Amanda\Uni\Datathon 2024\Datathon2024_heybanco.xlsm')
df_filtered = delete_comments(df, 'gracias') # eliminar comentarios con gracias
df_filtered = delete_comments(df, 'nuestros fans') # eliminar comentarios con gracias
df_filtered = delete_comments(df, str('+23K')) # eliminar comentarios con gracias


comments = df_filtered.iloc[:,2] #CAMBIAR A COMMENTS

filtered_sentences = []
for i in range(len(comments)):
    sentence = df_filtered.iloc[i,2]
    words = sentence.split() # dividir las palabras por espacios y coma
    words_to_remove = ['de', 'la','el','que','lo','en','me','por','nos','y','a','ya','con','para','mi','es','gracias','un','los','una','las','se','te','me', 'muchas','heybanco'] # palabras que queremos eliminar   
    words = [word.lower() for word in words] # poner las palabras en minusculas
    words = [word for word in words if word not in words_to_remove] # quitar palabras no deseadas
    words = [word.replace("!", "").replace("?", "").replace(".", "").replace("¡", "").replace(",", "").replace("-", "") for word in words] # eliminar signos de exclamacio
    words_plus = ' '.join(words) # convertir lista a string
    filtered_sentences.append(words_plus)


# # CONTAR PALABRAS
# #words = [word.lower() for word in words] # convertir las palabras a minúsculas
# #words = [word.replace("!", "").replace("?", "") for word in words] # eliminar signos de exclamacion
# words_to_remove = ['de', 'la','el','que','lo','en','me','por','nos','y','a','ya','con','para','mi','es','gracias','un','los','una','las','se','te','me', 'muchas', 'gracias','heybanco', 'hey','banco','como','hola','pero', 'les','sorprender','nuestros fans']
# print("hola")
# # Remove specific words from the list
# for i in 1:len(comments):
#     for sublist in words:
#         words = [word for word in words if word not in words_to_remove]  # Remove specific words
    
#     #reversed_list = [' '.join(sublist) for sublist in words_plus]
#     # Join the remaining words back into a phrase
# 
# print(words)
# # Now you can work with your DataFrame
# #word_counts = Counter(words)

#VECTORIZACION
#Generate trigrams
tokens = [word_tokenize(word) for word in filtered_sentences]

# Generate trigrams for each comment
trigrams_per_comment = [list(ngrams(token, 5)) for token in tokens]

# Flatten the list of trigrams
all_trigrams = [trigram for sublist in trigrams_per_comment for trigram in sublist]
# Count the frequency of trigrams
trigram_counts = Counter(all_trigrams)

# Get the 5 most common trigrams
most_common_trigrams = trigram_counts.most_common(10)

# Output the 5 most common trigrams
print(most_common_trigrams)