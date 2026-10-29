# utils/mnemonic.py

import random
from keybert import KeyBERT

# Easy-word dictionary for each alphabet letter
easy_words = {
    'A': ["apple", "ant", "arm", "arrow", "animal", "art", "air", "axe","aeroplane","all","alter","axe"],
    'B': ["ball", "bat", "bag", "bee", "bird", "boat", "book", "box", "bread", "bowl", "brush", "bush"],
    'C': ["cat", "car", "cup", "cap", "cow", "cake", "coin", "coat", "cane", "candle", "clock", "cloud"],
    'D': ["dog", "door", "duck", "doll", "drum", "dish", "dust", "desk", "dirt", "dive", "dawn", "dare"],
    'E': ["egg", "ear", "eye", "edge", "evil","east", "echo", "eagle", "earth", "easy", "eat", "end", "even"],
    'F': ["fan", "fish", "fig", "fog", "fox", "frog", "fun", "fire", "fork", "fence", "foot", "farm", "face"],
    'G': ["gun", "gap", "gem", "goat", "girl", "gold", "game", "gift", "golf", "gown", "grape", "grass", "glue"],
    'H': ["hat", "hen", "hop", "hot", "hut", "hill", "hand", "hope", "home", "hive", "honey", "hair", "heart"],
    'I': ["ink", "ice", "ill", "ivy", "iron", "idea", "item", "inch", "island", "item", "invite", "image", "input"],
    'J': ["jam", "jar", "jet", "job", "joy", "jug", "jog", "join", "joke", "jewel", "jump", "jungle", "juice"],
    'K': ["key", "kit", "kid", "kite", "king", "kind", "kick", "keep", "knee", "knot", "knock", "know", "kitchen"],
    'L': ["lip", "log", "leg", "lid", "lap", "lamp", "leaf", "line", "link", "lamb", "lake", "lamp", "loud", "love"],
    'M': ["man", "map", "mat", "mix", "mud", "milk", "moon", "mail", "mop", "mug", "mask", "march", "mouse", "music"],
    'N': ["net", "nap", "nut", "nail", "name", "nose", "nest", "night", "note", "nurse", "new", "next", "noble", "number"],
    'O': ["owl", "oil", "oak", "odd", "open", "ox", "one", "orb", "oven", "ocean", "offer", "orange", "order", "orbit"],
    'P': ["pen", "pig", "pot", "pan", "pet", "pin", "pie", "pack", "park", "palm", "page", "piano", "pipe", "puzzle"],
    'Q': ["queen", "quiz", "quit", "quill", "quick", "quiet", "quest", "quack", "quilt", "quote", "quarry", "quality", "quarter"],
    'R': ["rat", "rug", "run", "row", "red", "rip", "rice", "rock", "ring", "road", "rain", "rose", "root", "rush"],
    'S': ["sun", "sea", "set", "sit", "sap", "saw", "sip", "sock", "sand", "seal", "star", "stick", "step", "stone"],
    'T': ["top", "tip", "tap", "toy", "tin", "ten", "toe", "time", "tree", "tape", "tank", "task", "tail", "talk"],
    'U': ["urn", "use", "up", "us", "unit", "undo", "ugly", "upon", "under", "upset", "urban", "usual", "utensil", "ultimate"],
    'V': ["van", "vet", "vow", "vine", "view", "vent", "vote", "vast", "vase", "vocal", "village", "victory", "video", "vital"],
    'W': ["wet", "win", "was", "war", "way", "web", "wax", "wood", "wish", "wall", "wave", "wool", "work", "world"],
    'X': ["x-ray", "xmas", "xenon", "xerox", "xylem", "xiphoid", "xenial", "xenophobia", "xenolith", "xenogenesis", "xanthic", "xenon", "xylophone"],
    'Y': ["yam", "yes", "yet", "yarn", "yoga", "yard", "yell", "year", "yolk", "young", "youth", "yacht", "yellow", "yummy"],
    'Z': ["zip", "zoo", "zap", "zen", "zig", "zone", "zero", "zoom", "zest", "zinc", "zebra", "zeal", "zodiac", "zombie"]
}

def generate_mnemonic(initials):
    mnemonic_words = []
    for letter in initials:
        if letter in easy_words:
            mnemonic_words.append(random.choice(easy_words[letter]))
        else:
            mnemonic_words.append(letter)
    return ' '.join(mnemonic_words)

def extract_keywords(sentence, top_n=6):
    kw_model = KeyBERT()
    keywords = [kw[0] for kw in kw_model.extract_keywords(sentence, top_n=top_n)]
    return keywords

def create_mnemonic(sentence, top_n=6):
    keywords = extract_keywords(sentence, top_n=top_n)
    initials = [word[0].upper() for word in keywords]
    mnemonic_sentence = generate_mnemonic(initials)
    return keywords, initials, mnemonic_sentence
