import sys
import os
from typing import Text
import pandas as pd
import re
import nltk
import matplotlib.pyplot as plt
import custom
import wordcloud as wd
from nltk import word_tokenize
from nltk import FreqDist
from nltk.corpus import stopwords

# ----- GLOBAL VARIABLES ------

OUTPUT_CSV = './output.csv'

# ----- HELPER FUNCTIONS ------


def checkForOutput():
    print('Checking for "output.csv" and/or "output.xlsx"...')
    if not os.path.isfile(OUTPUT_CSV):
        print('File "output.csv" not found.')
        sys.exit()


def processFile():
    print('Processing "output.csv..."')
    df = pd.read_csv(OUTPUT_CSV)
    print('Eliminating possible duplicates...')
    df = df.drop_duplicates(subset='Team Request Id', keep="first")
    l = []

    for i in df.index:
        l.append(df['Role Notes'][i])

    l = [' '.join(l)]
    l = str(l)
    l = re.sub(r'\\n', ' ', l)

    return l


def processText(text):
    print('Processing text...')
    words = word_tokenize(text)

    words_no_punc = []
    for w in words:
        if w.isalpha():
            words_no_punc.append(w.lower())

    global stopwords
    stopwords = stopwords.words('dutch') + stopwords.words('english')

    clean_words = []
    for w in words_no_punc:
        if w not in stopwords:
            clean_words.append(w)

    modified_words = []
    for w in clean_words:
        if w not in custom.wordlist:
            modified_words.append(w)

    return modified_words


def findFreqDist(l):
    fdist = FreqDist(l)
    return fdist


def printCommonWords(fdist):
    print('\n')
    common_words = fdist.most_common(NUM_WORDS)
    print('Top %d common words:' % NUM_WORDS)
    print('--------------------')
    for q, w in common_words:
        print(f"{q}: {w}")
    print('--------------------')
    print('\n')


def plotFrequencyGraph(fdist):
    print('Plotting frequency graph...')
    fdist.plot(NUM_WORDS)


def generateWordCloud(l):
    print('Generating word cloud...')
    l = [' '.join(l)]
    text = str(l)

    wordcloud = wd.WordCloud().generate(text)
    plt.figure(figsize=(12, 12))
    plt.imshow(wordcloud)

    plt.axis("off")
    plt.show()

# ----- MAIN ------


print('\n')
print('Word Analysis based on output.csv')
print('-------------------------------------------------------------------------')
print('\n')
print('!!! Make sure "output.csv" is in the same folder as this script.')
print('!!! Column "Role Notes" is aggregated, stopwords (NL/EN) and words in "custom.wordlist" are eliminated.')
print('!!! Type CTRL + C to break the program at any time.')
print('\n')

NUM_WORDS = int(input('Number of words you want to see in result: '))

checkForOutput()

file = processFile()

word_list = processText(file)

fdist = findFreqDist(word_list)

printCommonWords(fdist)

graph = input('Do you want to plot a graph? (y/n) ')
if graph == 'y':
    plotFrequencyGraph(fdist)

wc = input('Do you want to generate a word cloud? (y/n) ')
if wc == 'y':
    generateWordCloud(word_list)
