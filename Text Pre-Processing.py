import pandas as pd
import os
from nltk.corpus import stopwords
from nltk.stem import PorterStemmer
porter=PorterStemmer()

"""
A program that cleans and pre-processes a text document.

"""

# define CWD
os.chdir(r"C:\Users\chris\Documents\Transgola\Clients\PROJECTS\2022\463230322_TM_MSF\Orignals\MemoQOut") 

# choose language pairs
source = "English"
target = "Portuguese"

# define stopwords object
stop = stopwords.words(source.lower())

#text to be processed
df = pd.read_excel(r"MemoQOut.xlsx", names=[source, target])

# make sure everything (dates, numbers etc.) are strings.
df = df.astype('str')

# keep a record of pre-filtered strings
preFiltered = df[source].tolist()

# remove white spaces at begining and end of string
df[source] = df[source].apply(lambda x: x.strip())

# create normalised column to filter on
df[f'{source}_2'] =   (df[source]
                       .str.lower()
                       .str.replace(r'[^\w\s]', '', regex=True) #remove punctuation
                       .apply(lambda x: ' '.join([word for word in x.split() if len(word) > 2]))
                       .apply(lambda x: ' '.join([word for word in x.split() if word not in (stop)])) # remove stopwords
                       .apply(lambda x: ' '.join([word for word in x.split() if not any(c.isdigit() for c in word)])) # remove words that contain numbers
                       .apply(lambda x: ' '.join([porter.stem(word) for word in x.split()]))
                       .str.replace(r'\d+', '', regex=True) #remove number digits
                       .str.replace('/\s\s+/g', ' ', regex=True) # no double white space, newlines, tabs
                       )

# perform de-dup on specified columns only (normalised column)
df.drop_duplicates(subset=[f'{source}_2'], inplace=True, keep="first")
df.drop(columns = f'{source}_2', inplace = True)
df.reset_index(inplace = True, drop = True)

# filter out unwanted tags
df = (
    df
    #.replace('nan','')
    .replace(r'\[[0-9+]\]?','', regex = True)
    .replace(r'\[[0-9+]\}?','', regex = True)
    .replace(r'\{[0-9+]\]?','', regex = True)
    .replace(r'\}\}?','', regex = True)
    .replace(r'\]?|\[?','', regex = True)
    .replace("•", "", regex = True)
    .replace(r'\S+@\S+', "" , regex = True) # email addresses
    .replace("", "EmptyCell")    
)

df[source] = df[source].apply(lambda x: x.replace(r'-', "", 1) if x.startswith('-') else x)    
    
# filter cells
df = df[(df[source].str.contains("EmptyCell")==False)]
df = df[df[source].str.contains('[A-Za-z]')] # remove any string that doesn't contain letters
df = df[df[source].str.len()>1]

# extras
df = df[(~df[source].str.startswith(("Website", "E-mail", "http", "www.", "Facebook")))  
        | (df[source].str.endswith(("Cleaning) *1", "SARS-Cov-2")))]

# create list containing post-filtered elements
postFiltered = df[source].tolist()

# output cleaned (tags), dups removed, filtered version of MemoQ output
df.to_excel('memoQOutClean.xlsx', index = False)

# set character length for split
splitThreshold = 90

# define splits
forQuill =df[(df[source].str.len()>=splitThreshold) | (df[source].str.contains(",")) | (df[source].str.contains("\?"))]
forQuill =forQuill[(forQuill[target] == 'nan')]
forQuill.to_excel('forQuill.xlsx', index=False)

forDeepL =df[(df[source].str.len()<splitThreshold) & (~df[source].str.contains("\?")) & (~df[source].str.contains(","))]
forDeepL =forDeepL[(forDeepL[target] == 'nan')]
forDeepL.to_excel('forDeepL.xlsx', index=False)

# reporting
print(f"{len(preFiltered)} = Pre-filtered length")
print(f"{len(postFiltered)} = Post-filtered length")
