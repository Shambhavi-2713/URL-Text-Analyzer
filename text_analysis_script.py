import pandas as pd
import os
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
import string
from collections import Counter

nltk.download('punkt')
nltk.download('StopWords')

def load_stop_words(filepath):
    stop_words = set()
    with open(filepath, 'r', encoding='latin-1') as file:
        words = file.read().split()
        stop_words.update(words)
    return stop_words

custom_stop_words = set()
stopword_folder = 'StopWords'

for filename in os.listdir(stopword_folder):
    filepath = os.path.join(stopword_folder, filename)
    custom_stop_words.update(load_stop_words(filepath))

stop_words = set(stopwords.words('english')).union(custom_stop_words)
punctuation = set(string.punctuation)

def clean_and_tokenize(text):
    text = ''.join([ch for ch in text if ch not in punctuation])
    text = text.lower()
    tokens = word_tokenize(text)
    tokens = [word for word in tokens if word not in stop_words]
    return tokens

def calculate_sentiment_scores(tokens):
    positive_words = set()
    negative_words = set()
    master_dict_path = os.path.join(os.getcwd(), 'MasterDictionary')
    with open(os.path.join(master_dict_path, 'positive-words.txt'), 'r') as f:
        positive_words.update(word.strip() for word in f.readlines())
    with open(os.path.join(master_dict_path, 'negative-words.txt'), 'r') as f:
        negative_words.update(word.strip() for word in f.readlines())
    positive_score = sum(1 for word in tokens if word in positive_words)
    negative_score = sum(1 for word in tokens if word in negative_words)
    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(tokens) + 0.000001)
    return positive_score, negative_score, polarity_score, subjectivity_score

def calculate_readability(text):
    sentences = sent_tokenize(text)
    words = clean_and_tokenize(text)
    avg_sentence_length = len(words) / len(sentences) if len(sentences) > 0 else 0
    complex_words = [word for word in words if syllable_count(word) > 2]
    percentage_complex_words = (len(complex_words) / len(words)) * 100 if len(words) > 0 else 0
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    avg_words_per_sentence = len(words) / len(sentences) if len(sentences) > 0 else 0
    return avg_sentence_length, percentage_complex_words, fog_index, avg_words_per_sentence

def syllable_count(word):
    vowels = 'aeiouy'
    count = 0
    last_char_was_vowel = False
    for char in word:
        if char.lower() in vowels:
            if not last_char_was_vowel:
                count += 1
            last_char_was_vowel = True
        else:
            last_char_was_vowel = False
    if word.endswith('e'):
        count -= 1
    if word.endswith('le'):
        count += 1
    if count == 0:
        count = 1
    return count

def calculate_other_metrics(text):
    words = clean_and_tokenize(text)
    word_count = len(words)
    syllables = sum(syllable_count(word) for word in words)
    avg_word_length = sum(len(word) for word in words) / len(words) if len(words) > 0 else 0
    personal_pronouns = ['i', 'we', 'my', 'ours', 'us']
    personal_pronoun_count = sum(1 for word in words if word.lower() in personal_pronouns)
    return word_count, syllables, avg_word_length, personal_pronoun_count

df = pd.read_excel('Input.xlsx')

output_data = []

for index, row in df.iterrows():
    url_id = row['URL_ID']
    article_file = f'articles/{url_id}.txt'
    try:
        with open(article_file, 'r', encoding='utf-8') as file:
            article_text = file.read()
    except FileNotFoundError:
        print(f'Article {url_id} not found. Skipping...')
        continue
    tokens = clean_and_tokenize(article_text)
    positive_score, negative_score, polarity_score, subjectivity_score = calculate_sentiment_scores(tokens)
    avg_sentence_length, percentage_complex_words, fog_index, avg_words_per_sentence = calculate_readability(article_text)
    word_count, syllables, avg_word_length, personal_pronoun_count = calculate_other_metrics(article_text)
    output_row = {
        'URL_ID': url_id,
        'Positive Score': positive_score,
        'Negative Score': negative_score,
        'Polarity Score': polarity_score,
        'Subjectivity Score': subjectivity_score,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
        'COMPLEX WORD COUNT': len([word for word in tokens if syllable_count(word) > 2]),
        'WORD COUNT': word_count,
        'SYLLABLE PER WORD': syllables / word_count if word_count > 0 else 0,
        'PERSONAL PRONOUNS': personal_pronoun_count,
        'AVG WORD LENGTH': avg_word_length
    }
    output_data.append(output_row)

output_df = pd.DataFrame(output_data)
output_file = 'Output Data Structure.xlsx'
output_df.to_excel(output_file, index=False)

print(f'Output saved to {output_file}')