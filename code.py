import spacy
import openpyxl
from pathlib import Path
import numpy as np
import re

nlp = spacy.load('en_core_web_sm')

url_regex = 'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
mentions_regex = '@\S+'
htmlTags_regex = '\<[/]?\S+\>'
hashtags_regex = '#\S+'
unicode_regex = '[^a-z^\ ]'
spaces_regex = '\s+'

#read text present in xlsx file. returns an array with all the tweets.
def read_file(file_name):
  file_path = Path(file_name)
  workbook = openpyxl.load_workbook(filename=file_path, read_only=True)
  sheet = workbook.active
  text = []
  for row in sheet.iter_rows(max_col=1):
      for cell in row:
          text.append(cell.value)
  return text

#Given an array containing text returns a sorted numpy array with every word
#present in the texts of the given array. Words will be present once despite there
# being multiple instances of said word.
#Punctuators, english stopwords, urls, hastags, numbers, unicode characters and spaces will be filtered.
#The stopwords used are the ones defined in spacy library.
def get_vocabulary_from_text(text):
  tokens = []
  for row in text:
    doc = nlp(row)
    for token in doc:
      word = token.text
      url_filter = re.findall(url_regex, word)
      mentions_filter = re.findall(mentions_regex, word)
      hashtags_filter = re.findall(hashtags_regex, word)
      unicode_filter = re.findall(unicode_regex, word)
      spaces_filter = re.findall(spaces_regex, word)
      if not token.is_stop and not token.is_punct and not url_filter and not mentions_filter and not hashtags_filter  and not unicode_filter and not spaces_filter and not word.isdigit():
        if(word.islower()):
          tokens.append(word)
        else:
          tokens.append(word.lower())

  wordList = np.array(tokens)
  uniqueTokens = np.unique(wordList)
  uniqueTokens.sort()
  return uniqueTokens

text = read_file('COV_train.xlsx')
uniqueTokens = get_vocabulary_from_text(text)

file = open("vocabulario.txt", "w+")
file.write("Number of tokens: " + str(uniqueTokens.size) + "\n")
for word in uniqueTokens:
  file.write(word + "\n")
file.close()