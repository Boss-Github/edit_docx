from docx import Document
from docx.shared import Pt

def replace_text_in_docx(file_path, placeholders, replacements, font=None, font_size=None):
    try:
        doc = Document(file_path)
        for placeholder, replacement in zip(placeholders, replacements):
            for paragraph in doc.paragraphs:
                if placeholder in paragraph.text:
                      for run in paragraph.runs:
                        if placeholder in run.text:
                            # Preserve font and font size if provided
                            if font:
                                run.font.name = font
                            if font_size:
                                run.font.size = Pt(font_size)
                            # Replace text
                            run.text = run.text.replace(placeholder, replacement)
            # Save as a new file with the replacement word in the file name
            new_file_path = file_path.replace('.docx', f'_{replacement}.docx')
            doc.save(new_file_path)
            print(f"Text replaced successfully. New file created: {new_file_path}")
    except Exception as e:
        print("An error occurred:", e)

# Example usage
file_path = "paradigma.docx"  # Replace with your file path
placeholders = ["!!__!!", "##__##", "{{__}}"]  # Replace with your placeholders
replacements = ["new_word_1", "new_word_2", "new_word_3"]  # Replace with replacement words
font = "Times New Roman"  # Replace with the font name you want
font_size = 16  # Replace with the font size you want
replace_text_in_docx(file_path, placeholders, replacements, font, font_size)
