import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches

folder_path = 'case'
for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)

paragraph.paragraph_format.line_spacing = 1.5
paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  

doc.save(doc_path)