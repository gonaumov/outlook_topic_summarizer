import json
import re

import nltk
from nltk.corpus import stopwords
from nltk.stem.wordnet import WordNetLemmatizer
import string
from gensim import corpora
from gensim.models.ldamodel import LdaModel

nltk.download('stopwords')
nltk.download('wordnet')

stop = set(stopwords.words('english'))
exclude = set(string.punctuation)
lemma = WordNetLemmatizer()


def clean(doc):
    stop_free = " ".join([i for i in doc.lower().split() if i not in stop])
    punc_free = " ".join([ch for ch in stop_free if ch not in exclude])
    return punc_free


def get_body_properties(filename):
    with open(filename, encoding='utf-16') as f:
        data = json.loads(f.read(), strict=False)
        return [item['body'] for item in data if 'body' in item]


def main():
    doc_complete = get_body_properties('emails.json')
    doc_clean = [map(lambda i: re.sub(r'\s+', ' ', i), clean(doc).split('   ')) for doc in doc_complete]
    dictionary = corpora.Dictionary(doc_clean)
    doc_term_matrix = [dictionary.doc2bow(doc) for doc in doc_clean]
    Lda = LdaModel
    ldamodel = Lda(doc_term_matrix, num_topics=12, id2word=dictionary, passes=90)
    print(ldamodel.show_topics(num_topics=100, num_words=20, log=False, formatted=True))


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
