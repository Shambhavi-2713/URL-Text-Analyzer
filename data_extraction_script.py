import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

input_file = 'Input.xlsx'
df = pd.read_excel(input_file)

if not os.path.exists('articles'):
    os.makedirs('articles')

def extract_text(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        title = soup.find('h1').get_text()
        paragraphs = soup.find_all('p')
        article_text = ' '.join([p.get_text() for p in paragraphs])
        return title, article_text
    except Exception as e:
        print(f'Error extracting text from {url}: {e}')
        return None, None

for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    title, article_text = extract_text(url)
    if article_text:
        with open(f'articles/{url_id}.txt', 'w', encoding='utf-8') as file:
            file.write(f'{title}\n\n{article_text}')
        print(f'Successfully saved article {url_id}')
    else:
        print(f'Failed to extract article {url_id}')
