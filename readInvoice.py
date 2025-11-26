import os
from typing import List
from pypdf import PdfReader

def collect_pdfs(path: str) -> List[str]:
    if os.path.isdir(path):
        return sorted(
            [
                os.path.join(path, f)
                for f in os.listdir(path)
                if f.lower().endswith(".pdf")
            ]
        )
    return [path]

def read_pdf(input_path: str) -> PdfReader:
    if not os.path.exists(input_path):
        raise FileNotFoundError(input_path)
    if not input_path.lower().endswith(".pdf"):
        raise ValueError("not a pdf")
    return PdfReader(input_path)