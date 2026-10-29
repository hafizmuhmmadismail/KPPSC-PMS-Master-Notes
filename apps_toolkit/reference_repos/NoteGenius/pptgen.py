# utils/pptgen.py

from pptx import Presentation
from pptx.util import Pt
import textwrap
import os
import datetime

def generate_ppt(input_text):
    # Split text into sentences
    sentences = [s.strip() for s in input_text.replace("\n", " ").split(".") if s.strip()]

    # Group sentences into chunks (5 sentences per slide)
    chunk_size = 5
    sections = [" ".join(sentences[i:i+chunk_size]) for i in range(0, len(sentences), chunk_size)]

    prs = Presentation()

    for idx, section in enumerate(sections, 1):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"Slide {idx}"

        # Get content placeholder
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()

        # Add sentences as bullets
        for sentence in section.split("."):
            sentence = sentence.strip()
            if sentence:
                p = tf.add_paragraph()
                p.text = sentence + "."
                p.level = 0
                p.font.size = Pt(18)

    # Save with timestamp to avoid overwrite
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"pptx_output_{timestamp}.pptx"
    output_path = os.path.join("instance", filename)
    prs.save(output_path)

    return output_path
