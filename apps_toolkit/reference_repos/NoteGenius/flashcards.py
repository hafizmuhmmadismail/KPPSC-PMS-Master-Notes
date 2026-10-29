# utils/flashcards.py

from rake_nltk import Rake
import nltk
from nltk.tokenize import sent_tokenize
from rapidfuzz import fuzz

# Download required NLTK data
nltk.download('stopwords')
nltk.download('punkt')

def generate_flashcards_from_keywords(text, num_keywords=5):
    """
    Generate flashcards based on keywords extracted from the text.
    """
    rake = Rake()
    rake.extract_keywords_from_text(text)
    keywords = rake.get_ranked_phrases()[:num_keywords]

    sentences = sent_tokenize(text)
    flashcards = []
    for keyword in keywords:
        for sentence in sentences:
            if fuzz.partial_ratio(keyword.lower(), sentence.lower()) > 80:
                flashcards.append({'keyword': keyword, 'context': sentence.strip()})
                break
    return flashcards
