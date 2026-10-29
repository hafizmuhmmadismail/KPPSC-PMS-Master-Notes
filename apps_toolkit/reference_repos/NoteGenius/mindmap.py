# utils/mindmap.py

import spacy
import re
from graphviz import Digraph

# Load spaCy model
nlp = spacy.load("en_core_web_sm")

def split_sentence(sentence):
    """
    Split a sentence into main clause and sub-points.
    """
    clauses = re.split(r",| and | including | like | through ", sentence)
    clauses = [c.strip() for c in clauses if c.strip()]
    if len(clauses) > 1:
        main = clauses[0]
        subs = clauses[1:]
    else:
        main = clauses[0]
        subs = []
    return main, subs


def generate_mindmap(text, filename="mindmap", fmt="png"):
    """
    Generate a mindmap diagram from the input text.
    Saves it as a PNG file.
    Returns the output filename.
    """
    # Step 1: Split text into sentences
    doc = nlp(text)
    sentences = [sent.text.strip() for sent in doc.sents]

    # Step 2: Build mindmap dictionary
    mind_map_dict = {}
    for sent in sentences:
        main, subs = split_sentence(sent)
        mind_map_dict[main] = subs

    # Step 3: Extract title automatically
    title = ""
    for ent in doc.ents:
        if ent.label_ in ["PERSON", "ORG"]:
            title = ent.text
            break
    if title == "":
        title = sentences[0].split(",")[0]

    # Step 4: Generate Graphviz diagram
    g = Digraph('G', filename=filename, format=fmt)
    g.attr(rankdir='LR', splines='spline', pad='0.3', nodesep='0.4', ranksep='1')
    g.attr('node', shape='box', style='rounded,filled', fontname='Helvetica', fontsize='14')
    g.attr('edge', arrowsize='0.7', penwidth='1.2')

    # Central node
    g.node('CENTER', title, fillcolor='#0E2B50', fontcolor='white', fontsize='18', penwidth='0')

    colors = ["#9EE2F7", "#F6A74B", "#F59AC5", "#77D587", "#E1A1F7", "#FFD580"]

    for i, (branch, points) in enumerate(mind_map_dict.items()):
        branch_id = f"B{i}"
        color = colors[i % len(colors)]
        g.node(branch_id, branch, fillcolor=color, penwidth='0')
        g.edge('CENTER', branch_id)

        for j, point in enumerate(points):
            child_id = f"B{i}C{j}"
            g.node(child_id, point, fillcolor="#FFFFFF", style="rounded", penwidth='1', color="#BFBFBF")
            g.edge(branch_id, child_id)

    output_file = g.render(cleanup=True)
    return output_file
