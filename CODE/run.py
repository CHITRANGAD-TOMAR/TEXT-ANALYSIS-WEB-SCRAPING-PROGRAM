import subprocess
import sys
import requests
import os
import pandas as pd
import os
import re
import nltk
import time
import pandas as pd

from nltk.tokenize import word_tokenize
from nltk.corpus import cmudict
from bs4 import BeautifulSoup
from  colorama import Fore, Style, init


script_path = os.path.abspath(__file__)
cwd = os.path.dirname(script_path)

nltk.download('punkt')
nltk.download('cmudict')
d = cmudict.dict()


required_packages = [
    'nltk',
    'pandas',
    'request'
    'bs4'
    'beautifulsoup4'
    're'
    'colorama'
]


"""
Installs the required dependencies
"""

def INSTALL_PACKAGES(packages):
    try:
        for package in packages:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"Successfully installed {package}")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install {package}. Error: {e}")


print(Fore.GREEN +"\nRequired dependencies installation started!\n"+ Style.RESET_ALL)
time.sleep(1)
INSTALL_PACKAGES(required_packages)
print(Fore.GREEN +"\nPackages installation completed!\n"+ Style.RESET_ALL)
time.sleep(1)
os.system('cls')


"""
Data extraction and saving to a text file
"""

def DATA_EXTRACTION():

    def extract_urls(file_path, sheet_name='Sheet1'):
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            url_ids = df["URL_ID"].tolist()
            urls = df["URL"].tolist()
            return url_ids,urls
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return []
    
    def extract_article(url):
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
        response = requests.get(url, headers=headers)    
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            article_title = soup.find('h1', class_="entry-title")
            article_content = soup.find('div', class_="td-post-content tagdiv-type")  
            try:
                article_title = article_title.get_text()
                article_text = article_content.get_text()
                return article_title,article_text
            except:
                print(Fore.RED +"Couldn't find article content on the page."+ Style.RESET_ALL)
                article_text = []
                article_title = soup.find_all('h1')
                for article_title in article_title:
                    article_title = article_title.text
                text = soup.find_all('p')
                for article in text:
                    article_text.append(article.text)
                article_text = ' '.join(article_text)
                return article_title,article_text
        else:
            print(Fore.RED +f"Failed to retrieve content from {url}. Status code: {response.status_code}"+ Style.RESET_ALL)
            article_title= ""
            article_text= ""
            return article_title,article_text
    
    def save_to_txt_file(text, filename):
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(text)

    file_path = cwd + '\\input.xlsx'
    url_ids,urls = extract_urls(file_path)
    url_dict = dict(zip(url_ids, urls))
    for idx, url in enumerate(urls):
        print(f"Processing URL {idx + 1}/{len(urls)}: {url}")
        article_title,article_text = extract_article(url)
        text = article_title + "\n" + article_text
        if article_title and article_text:
            extracted_data_path = cwd + f"\\EXTRACTED_DATA"
            if not os.path.exists(extracted_data_path):
                os.makedirs(extracted_data_path)
                filename = cwd + f"\\EXTRACTED_DATA\\{url_ids[idx]}.txt"
                save_to_txt_file(article_text, filename)
                print(f"Article extracted and saved to {filename}")
            else:    
                filename = cwd + f"\\EXTRACTED_DATA\\{url_ids[idx]}.txt"
                save_to_txt_file(article_text, filename)
                print(f"Article extracted and saved to {filename}")
        else:
            print(Fore.RED +"Failed to extract article from the URL."+ Style.RESET_ALL)
    return url_dict


print(Fore.GREEN +"\nData Extraction started!\n"+ Style.RESET_ALL)
time.sleep(1)
url_dict = DATA_EXTRACTION()
print(Fore.GREEN +"\nData Extraction completed!\n"+ Style.RESET_ALL)
time.sleep(1)
os.system("cls")    


"""
Text analysis and saving results to csv file
"""


def TEXT_ANALYSIS(url_dict):

    def read_stopwords(stopword_files):
        stopwords = set()
        for file in stopword_files:
            with open(file, 'r') as f:
                stopwords.update(f.read().split())
        return stopwords

    def remove_stopwords(text, stopwords):
        words = text.split()
        filtered_words = [word for word in words if word.lower() not in stopwords]
        return ' '.join(filtered_words)

    def process_text_file(input_file, stopwords):
        with open(input_file, 'r' , encoding='utf-8') as f:
            text = f.read()
        filtered_text = remove_stopwords(text, stopwords)
        return filtered_text

    def read_master_dictionary(folder_path):
        positive_words = []
        negative_words = []
        positive_file = os.path.join(folder_path, 'positive-words.txt') 
        negative_file = os.path.join(folder_path, 'negative-words.txt') 
        with open(positive_file, 'r') as f:
            for line in f:
                word = line.strip()
                if word:
                    positive_words.append(word)
        with open(negative_file, 'r') as f:
            for line in f:
                word = line.strip()
                if word:
                    negative_words.append(word)    
        return positive_words, negative_words
    
    def filter_positive_tokens(filtered_text, positive_words):
        tokens = word_tokenize(filtered_text)
        filtered_tokens = [word for word in tokens if word.lower() in positive_words]
        return filtered_tokens

    def filter_negative_tokens(filtered_text, negative_words):
        tokens = word_tokenize(filtered_text)
        filtered_tokens = [word for word in tokens if word.lower() in negative_words]
        return filtered_tokens

    def count_syllables(word):
        if word.lower() in d:
            return [len(list(y for y in x if y[-1].isdigit())) for x in d[word.lower()]][0]
        else:
            return 0

    def count_complex_words(text):
        words = text.split()
        complex_words = [word for word in words if count_syllables(word) >= 3]
        return complex_words

    def count_syllables(word):
        vowel_pattern = r'[aeiouyAEIOUY]'
        word = re.sub(r'(es|ed)$', '', word)
        syllable_count = len(re.findall(vowel_pattern, word))
        return max(syllable_count, 1)

    def syllable_count_per_word(text):
        words = re.findall(r'\b\w+\b', text)
        syllable_counts = {word: count_syllables(word) for word in words}
        return syllable_counts

    def count_personal_pronouns(text):
        pronoun_pattern = r'\b(I|we|my|ours|us)\b'
        matches = re.findall(pronoun_pattern, text, re.IGNORECASE)
        pronoun_counts = {pronoun.lower(): 0 for pronoun in ['I', 'we', 'my', 'ours', 'us']}
        for match in matches:
            pronoun_counts[match.lower()] += 1
        return pronoun_counts

    def average_word_length(text):
        words = text.split()
        total_characters = sum(len(word) for word in words)
        total_words = len(words)
        if total_words != 0:
            average_length = total_characters / total_words
        else:
            average_length = 0        
        return average_length
    
    def average_sentence_length(text):
        sentences = nltk.sent_tokenize(text)    
        total_words = 0
        total_sentences = len(sentences)
        for sentence in sentences:
            words = nltk.word_tokenize(sentence)
            total_words += len(words)
        if total_sentences > 0:
            average_length = total_words / total_sentences
            return average_length
        else:
            return 0


    def analysis(input_file,stop_words_folder,master_dictionary_folder): 
        stopword_files = []
        files_in_folder = os.listdir(stop_words_folder)
        for file_name in files_in_folder:
            stopword_files.append(stop_words_folder+"\\"+file_name)
        stopwords = read_stopwords(stopword_files)
        filtered_text = process_text_file(input_file, stopwords)
        positive_words, negative_words = read_master_dictionary(master_dictionary_folder)
        filtered_positive_tokens = filter_positive_tokens(filtered_text, positive_words)
        filtered_negative_tokens = filter_negative_tokens(filtered_text, negative_words)
        Total_Words = len(filtered_text.split())
        Positive_Score = float(len(filtered_positive_tokens))
        Negative_Score = float(len(filtered_negative_tokens))
        Polarity_Score = (Positive_Score - Negative_Score) / ((Positive_Score + Negative_Score) + 0.000001)
        Subjectivity_Score = (Positive_Score + Negative_Score)/ ((Total_Words) + 0.000001)
        
        sentences = filtered_text.split('. ')
        sentences.extend(filtered_text.split('! '))
        sentences.extend(filtered_text.split('? '))
        sentences = [sentence for sentence in sentences if sentence]
        number_of_sentences = len(sentences)
        number_of_words = len(filtered_text.split())

        Average_Sentence_Length = average_sentence_length(filtered_text)

        complex_words_list_count = count_complex_words(filtered_text)
        complex_words_list_count = len(complex_words_list_count)
        Percentage_of_Complex_words = complex_words_list_count / number_of_words
        Fog_Index = 0.4 * (Average_Sentence_Length + Percentage_of_Complex_words)
        Average_Number_of_Words_Per_Sentence = number_of_words / number_of_sentences
        syllable_counts = syllable_count_per_word(filtered_text)
        pronoun_counts = count_personal_pronouns(filtered_text)
        avg_word_length = average_word_length(filtered_text)
        return Positive_Score, Negative_Score, Polarity_Score, Subjectivity_Score, Average_Sentence_Length, Percentage_of_Complex_words, Fog_Index, Average_Number_of_Words_Per_Sentence, complex_words_list_count, number_of_words, syllable_counts, pronoun_counts, avg_word_length

    def CSV_WRITTING(file_name,url_id,url,Positive_Score, Negative_Score, Polarity_Score, Subjectivity_Score, Average_Sentence_Length, Percentage_of_Complex_words, Fog_Index, Average_Number_of_Words_Per_Sentence, count_complex_words, number_of_words, syllable_count_per_word, pronoun_counts, average_word_length):
        columns = [
            "URL_ID", "URL", "POSITIVE SCORE", "NEGATIVE SCORE", "POLARITY SCORE",
            "SUBJECTIVITY SCORE", "AVG SENTENCE LENGTH", "PERCENTAGE OF COMPLEX WORDS",
            "FOG INDEX", "AVG NUMBER OF WORDS PER SENTENCE", "COMPLEX WORD COUNT",
            "WORD COUNT", "SYLLABLE PER WORD", "PERSONAL PRONOUNS", "AVG WORD LENGTH" ]

        data = {
        "URL_ID": url_id,
        "URL": url,
        "POSITIVE SCORE": Positive_Score,
        "NEGATIVE SCORE": Negative_Score,
        "POLARITY SCORE": Polarity_Score,
        "SUBJECTIVITY SCORE": Subjectivity_Score,
        "AVG SENTENCE LENGTH": Average_Sentence_Length,
        "PERCENTAGE OF COMPLEX WORDS": Percentage_of_Complex_words,
        "FOG INDEX": Fog_Index,
        "AVG NUMBER OF WORDS PER SENTENCE": Average_Number_of_Words_Per_Sentence,
        "COMPLEX WORD COUNT": count_complex_words,
        "WORD COUNT": number_of_words,
        "SYLLABLE PER WORD": syllable_count_per_word,
        "PERSONAL PRONOUNS": pronoun_counts,
        "AVG WORD LENGTH": average_word_length }
        
        df = pd.DataFrame(columns=columns)
        new_data = pd.DataFrame([data])
        if os.path.isfile(file_name):
            new_data.to_csv(file_name, mode='a', header=False, index=False)
        else:
            new_data.to_csv(file_name, mode='w', header=columns, index=False)


    stop_words_folder = cwd + "\\StopWords"  
    master_dictionary_folder = cwd + "\\MasterDictionary"
    extracted_data_folder = cwd + "\\EXTRACTED_DATA"
    output_filename = cwd + "\\output.csv"
    
    files_in_folder = os.listdir(extracted_data_folder)
    
    for file_name in files_in_folder:
        input_file = extracted_data_folder+"\\"+file_name
        Positive_Score, Negative_Score, Polarity_Score, Subjectivity_Score, Average_Sentence_Length, Percentage_of_Complex_words, Fog_Index, Average_Number_of_Words_Per_Sentence, complex_words_list_count, number_of_words, syllable_counts, pronoun_counts, avg_word_length = analysis(input_file,stop_words_folder,master_dictionary_folder)
        
        urlid = os.path.splitext(os.path.basename(file_name))[0]
        if urlid in url_dict:
            url =  url_dict[urlid]
        else:
            url = ""
            urlid = ""

        CSV_WRITTING(output_filename,urlid,url,Positive_Score, Negative_Score, Polarity_Score, Subjectivity_Score, Average_Sentence_Length, Percentage_of_Complex_words, Fog_Index, Average_Number_of_Words_Per_Sentence, complex_words_list_count, number_of_words, syllable_counts, pronoun_counts, avg_word_length)
    
    print("data saved successfully in output.csv file!\n")


print(Fore.GREEN +"\nText Analysis in progress!\n"+ Style.RESET_ALL)
time.sleep(1)
TEXT_ANALYSIS(url_dict)
print(Fore.GREEN +"\nText Analysis completed!\n"+ Style.RESET_ALL)
time.sleep(1)
