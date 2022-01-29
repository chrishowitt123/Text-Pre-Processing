import pandas as pd
import os
os.chdir(r"C:\Users\chris\Documents\Transgola\Clients\PROJECTS\2022\445280122_TM_MSF\Orignals\MemoQ") 

# choose language pairs
source = "English"
target = "Portuguese"

df = pd.read_excel(r"MemoQOut.xlsx", names=[source, target])

# make sure everything (dates, numbers etc.) are strings.
df = df.astype('str')

preFiltered = df[source].tolist()

# remove white spaces at begining and end of string
df[source] = df[source].apply(lambda x: x.strip())

# create normalised column to filter on (lower)
df[f'{source}_2'] = df[source].str.lower().str.rstrip("'°0123456789.-:;")

# subset performs operation on specified columns only (normalised column)
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
    .replace(r"(Table|Figure) [0-9]+:", "", regex = True) 
    .replace("", "EmptyCell")
    
)

# create list containing pre-filtered elements

# filter cells
df = df[(df[source].str.contains("EmptyCell")==False)]
df = df[df[source].str.contains('[A-Za-z]')] 
df = df[df[source].str.len()>1]

# filter out strings ending with numbers but keep anything relivant
df = df[(~df[source].str.startswith(("Website", "E-mail", "http:", "www.", "Facebook")))   #(~df[source].str.endswith(("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
        | (df[source].str.endswith(("Cleaning) *1", "SARS-Cov-2")))]

# create list containing post-filtered elements
postFiltered = df[source].tolist()

# output cleaned (tags), dups removed, filtered version of MemoQ output
df.to_excel('memoQOutClean.xlsx')

# set character length for split
splitThreshold = 90

#define splits
forQuill =df[(df[source].str.len()>=splitThreshold) | (df[source].str.contains(",")) | (df[source].str.contains("\?"))]
forQuill =forQuill[(forQuill[target] == 'nan')]
forQuill.to_excel('forQuill.xlsx', index=False)

forDeepL =df[(df[source].str.len()<splitThreshold) & (~df[source].str.contains("\?")) & (~df[source].str.contains(","))]
forDeepL =forDeepL[(forDeepL[target] == 'nan')]
forDeepL.to_excel('forDeepL.xlsx', index=False)

# check that the cleaned output has the same amount of rows as the splits combined

if len(df[source]) == len(forQuill[target]) + len(forDeepL[source]):
    print("MemoQ master and DeepL + Quill splits are of equal length!\n")
else:
    print("MemoQ master and DeepL + Quill splits are NOT of equal length!\n")

print(f"{len(preFiltered)} = Pre-filtered length")
print(f"{len(postFiltered)} = Post-filtered length")
print("\n")
print("Removed during filtering\n")

results = set(preFiltered) - set(postFiltered)
for r in sorted(results,  key=len, reverse=True):
    print(r)
