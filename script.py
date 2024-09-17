import spacy
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import Color
from reportlab.lib.units import inch

# Carica il modello NLP di spaCy
nlp = spacy.load('en_core_web_sm')

# Configura margini e interlinea
MARGINE_LEFT = 50
MARGINE_RIGHT = 50
MARGINE_TOP = 50
MARGINE_BOTTOM = 50
INTERLINEA = 15


def riconosci_struttura_testo(para):
    doc = nlp(para)

    if len(para.split()) <= 5 and para.istitle():
        return "title"
    elif para.startswith('"') or para.startswith("'"):
        return "quote"
    elif len(para.split()) <= 3:
        return "signature"
    elif "\n" in para or len(para) < 50:
        return "poem"
    else:
        return "paragraph"


def wrap_text(pdf, text, font_name, font_size, x, y, max_width):
    lines = []
    words = text.split()
    line = ""

    for word in words:
        test_line = line + " " + word if line else word
        test_width = pdf.stringWidth(test_line, font_name, font_size)

        if test_width > max_width:
            if line:  # Add the previous line
                lines.append(line)
            line = word  # Start a new line with the current word
        else:
            line = test_line

    if line:  # Add the last line
        lines.append(line)

    return lines


def crea_pdf_da_word_con_formattazione(word_file, output_pdf):
    doc = Document(word_file)
    pdf = canvas.Canvas(output_pdf, pagesize=A4)
    width, height = A4

    x = MARGINE_LEFT
    y = height - MARGINE_TOP
    max_width = width - MARGINE_LEFT - MARGINE_RIGHT

    for para in doc.paragraphs:
        tipo = riconosci_struttura_testo(para.text)

        if tipo in ["title", "quote", "signature"]:
            font_size = 18 if tipo == "title" else 14 if tipo == "quote" else 12
            font_name = "Helvetica-Bold" if tipo == "title" else "Helvetica-Oblique"
        else:
            font_size = 12
            font_name = "Helvetica"

        pdf.setFont(font_name, font_size)
        pdf.setFillColor(
            Color(0.2, 0.2, 0.8) if tipo == "title" else Color(0.3, 0.3, 0.3) if tipo == "quote" else Color(0, 0,
                                                                                                            0.6) if tipo == "signature" else Color(
                0, 0, 0))

        lines = wrap_text(pdf, para.text, font_name, font_size, x, y, max_width)

        for line in lines:
            if y - font_size < MARGINE_BOTTOM:
                pdf.showPage()
                pdf.setFont(font_name, font_size)
                y = height - MARGINE_TOP

            pdf.drawString(x, y, line)
            y -= INTERLINEA

        y -= INTERLINEA if tipo != "title" else INTERLINEA + 5

    pdf.save()


# Esempio di utilizzo
crea_pdf_da_word_con_formattazione(
    r"/Doc.docx",
    "output_formattato.pdf")
