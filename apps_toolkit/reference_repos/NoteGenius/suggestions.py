def suggest_links(input_text):
    # --- imports inside ---
    import requests
    from youtube_search import YoutubeSearch
    from sklearn.feature_extraction.text import TfidfVectorizer
    import numpy as np

    # --- keyword extraction ---
    vectorizer = TfidfVectorizer(ngram_range=(1, 2), stop_words='english')
    tfidf = vectorizer.fit_transform([input_text])
    scores = tfidf.toarray()[0]
    feature_names = np.array(vectorizer.get_feature_names_out())
    top_idx = scores.argsort()[-5:][::-1]
    keywords = feature_names[top_idx].tolist()
    query = " ".join(keywords) if keywords else input_text[:100]

    # --- Google Search API ---
    API_KEY = "AIzaSyAlgymwbmedb1sk_t1ttooJlynE_YCPQVQ"
    CSE_ID = "c661a8b45522242dd"
    google_links = []
    google_error = None

    if API_KEY.startswith("YOUR_") or CSE_ID.startswith("YOUR_"):
        google_error = "❌ Missing Google API credentials"
    else:
        search_url = "https://www.googleapis.com/customsearch/v1"
        params = {"key": API_KEY, "cx": CSE_ID, "q": query, "num": 5}
        try:
            resp = requests.get(search_url, params=params, timeout=10)
            data = resp.json()
            if "error" in data:
                google_error = f"❌ API error: {data['error'].get('message', 'Unknown error')}"
            else:
                items = data.get("items", []) or []
                google_links = [it.get("link", "") for it in items if it.get("link")]
        except Exception as e:
            google_error = f"❌ Request error: {e}"

    # --- YouTube Search ---
    videos = []
    try:
        results = YoutubeSearch(query, max_results=5).to_dict()
        for r in results:
            title = r.get("title", "Untitled")
            url = "https://www.youtube.com" + (r.get("url_suffix", "") or "")
            videos.append((title, url))
    except Exception as e:
        videos = [("Error fetching YouTube results", str(e))]

    # --- return everything ---
    return {
        "keywords": keywords,
        "query": query,
        "google": {"links": google_links, "error": google_error},
        "youtube": videos,
    }