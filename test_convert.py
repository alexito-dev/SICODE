from docx2pdf import convert
import pythoncom
import tempfile
import traceback

def test_convert():
    try:
        pythoncom.CoInitialize()
        template = "app/modulos/autodoc/plantillas/Plantilla_Base.docx"
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            pdf_path = tmp_pdf.name

        print("Intentando convertir...")
        convert(template, pdf_path)
        print("Convertido correctamente!")
    except Exception as e:
        print("Error details:")
        traceback.print_exc()

if __name__ == "__main__":
    test_convert()
