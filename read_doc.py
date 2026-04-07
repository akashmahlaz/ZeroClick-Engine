from docx import Document

doc = Document('In-content snippet (zero click)- overview (3).docx')

output = []
for para in doc.paragraphs:
    output.append(para.text)

print('\n'.join(output))