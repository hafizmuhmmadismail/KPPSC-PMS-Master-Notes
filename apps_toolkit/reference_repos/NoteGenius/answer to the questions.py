# utils/qna.py

from transformers import pipeline
import torch

# Auto-detect device: GPU if available, else CPU
device = 0 if torch.cuda.is_available() else -1

# Load the QA pipeline once when the module is imported
qa_pipeline = pipeline("question-answering", model="deepset/tinyroberta-squad2", device=device)

def get_answers(context, questions):
    results = []
    for question in questions:
        try:
            result = qa_pipeline(question=question, context=context)
            answer = result.get('answer', 'No answer found.')
        except Exception as e:
            answer = f"Error: {str(e)}"
        results.append({'question': question, 'answer': answer})
    return results
