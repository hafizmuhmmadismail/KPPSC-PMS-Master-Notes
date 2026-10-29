def ques_gen():
    
    import spacy
    import random
    import nltk
    from nltk.corpus import wordnet as wn
    from nltk.corpus import stopwords

    # ---------------- NLTK Setup ----------------
    try:
        nltk.data.find('corpora/wordnet')
        nltk.data.find('corpora/stopwords')
        nltk.data.find('corpora/omw-1.4')
    except LookupError:
        nltk.download('wordnet')
        nltk.download('stopwords')
        nltk.download('omw-1.4')

    # ---------------- spaCy Setup ----------------
    nlp = spacy.load("en_core_web_sm")
    stop_words = set(stopwords.words('english'))

    # ---------------- Question Generator ----------------
    class TextQuestionGenerator:
        def _init_(self):
            self.nlp = nlp
            self.stop_words = stop_words
            self.templates = [
                "Why is {keyword} important?",
                "What is the role of {keyword}?",
                "How does {keyword} work?",
                "Explain the significance of {keyword}.",
                "How has {keyword} impacted the field?"
            ]

        # --------- 1) Clean Sentences ---------
        def clean_sentences(self, text):
            doc = self.nlp(text)
            sentences = [sent.text.strip() for sent in doc.sents if 10 <= len(sent.text.strip()) <= 300]
            return sentences

        # --------- 2) Important Keywords ---------
        def important_keywords(self, sentence):
            doc = self.nlp(sentence)
            keywords = set()

            # Named Entities
            for ent in doc.ents:
                if ent.text.lower() not in self.stop_words and len(ent.text) > 2:
                    keywords.add(ent.text)

            # Noun Chunks
            for chunk in doc.noun_chunks:
                if chunk.root.pos_ in ["NOUN", "PROPN"] and chunk.text.lower() not in self.stop_words:
                    keywords.add(chunk.text)

            # Important adjectives and verbs
            for token in doc:
                if token.pos_ in ["ADJ", "VERB"] and token.text.lower() not in self.stop_words and len(token.text) > 3:
                    keywords.add(token.text)

            return list(keywords)

        # --------- 3) WordNet Helpers ---------
        def wordnet_synonyms(self, word, pos_tag=None):
            synonyms = set()
            synsets = wn.synsets(word.lower())
            if pos_tag:
                pos_map = {'NOUN':'n','VERB':'v','ADJ':'a','ADV':'r'}
                if pos_tag in pos_map:
                    synsets = [s for s in synsets if s.pos() == pos_map[pos_tag]]
            for syn in synsets:
                for lemma in syn.lemmas():
                    synonym = lemma.name().replace('_',' ')
                    if synonym.lower() != word.lower():
                        synonyms.add(synonym)
            return list(synonyms)

        def wordnet_antonyms(self, word):
            antonyms = set()
            for syn in wn.synsets(word.lower()):
                for lemma in syn.lemmas():
                    if lemma.antonyms():
                        for ant in lemma.antonyms():
                            antonyms.add(ant.name().replace('_',' '))
            return list(antonyms)

        def wordnet_hyponyms_hypernyms(self, word):
            related = set()
            for syn in wn.synsets(word.lower()):
                for hyponym in syn.hyponyms():
                    for lemma in hyponym.lemmas():
                        related.add(lemma.name().replace('_',' '))
                for hypernym in syn.hypernyms():
                    for lemma in hypernym.lemmas():
                        related.add(lemma.name().replace('_',' '))
            return list(related)

        # --------- 4) Context-based Filtering ---------
        def filter_keywords(self, keyword, sentence):
            """Filter out trivial or ambiguous keywords"""
            kw_lower = keyword.lower()
            if kw_lower in self.stop_words:
                return False
            if len(kw_lower) <= 2:
                return False
            if kw_lower in ["things","we","this","it","each","a"]:
                return False
            if kw_lower not in sentence.lower():
                return False
            return True

        # --------- 5) Generate Question ---------
        def generate_question(self, sentence):
            keywords = self.important_keywords(sentence)
            filtered = [kw for kw in keywords if self.filter_keywords(kw, sentence)]
            if not filtered:
                return None

            # Pick the longest keyword for better context
            keyword = max(filtered, key=len)
            template = random.choice(self.templates)
            question = template.format(keyword=keyword)
            return question

        # --------- 6) Process Document ---------
        def process_text(self, text):
            sentences = self.clean_sentences(text)
            questions = []
            for sent in sentences:
                q = self.generate_question(sent)
                if q:
                    questions.append(q)
            return questions

    # ---------------- Demo ----------------
    if _name_ == "_main_":
        text = """
        Solids, when heated, emit radiation varying over a wide range of wavelengths. For example: when we heat solid colour, changes continue with a further increase in temperature. This change in colour happens from a lower frequency region to a higher frequency region as the temperature increases. For example, in many cases, it changes from red to blue. An ideal body which can emit and absorb radiation of all frequencies is called a black body. The radiation emitted by such bodies is called black body radiation.

        Thus, we can say that variation of frequency for black body radiation depends on the temperature. At a given temperature, the intensity of radiation is found to increase with an increase in the wavelength of radiation which increases to a maximum value and then decreases with an increase in the wavelength. This phenomenon couldn't be explained with the help of Maxwell's suggestions. Hence, Planck proposed Planck's quantum theory to explain this phenomenon.
        
        Quantum theory revolutionized our understanding of physics. It introduced the concept that energy is quantized, meaning it comes in discrete packets called quanta. This theory helped explain many phenomena that classical physics couldn't, including the photoelectric effect and atomic spectra. The quantum mechanical model of the atom replaced the classical model and provided a more accurate description of atomic behavior.
        """
        tqg = TextQuestionGenerator()
        questions = tqg.process_text(text)
        
        print("\nGenerated Questions:")
        for i, q in enumerate(questions, 1):
            print(f"Q{i}: {q}")