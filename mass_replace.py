import os
import win32com.client
import pythoncom

def mass_replace() -> None:
    pythoncom.CoInitialize()
    plantillas_dir = r"D:\SICODE\APPS\AUTODOC\PLANTILLAS"

    # 1. Diccionario de reemplazos (Texto Actual -> Etiqueta Jinja2)
    replacements = {
        "0012-2025-SUNAT/7N0300": "{{ esquela }}",
        "03 de enero de 2025": "{{ fechaActual }}",
        "MUNICIPALIDAD DISTRITAL DE EL TAMBO": "{{ razonSocial }}",
        "20133696742": "{{ ruc }}",
        "AV. MARISCAL CASTILLA N°. 1920": "{{ domicilio }}",
        "EL TAMBO – HUANCAYO - JUNIN": "",  # Se borra porque el domicilio de arriba lo reemplazará
        "Solicitud de Reconocimiento de Pago con Error": "{{ subtipoPlantilla }}",
        "URD079-2024-1258677": "{{ numExpediente }}",
        "17.12.2024": "{{ fechaExpediente }}"
    }

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False  # Para que no abra ventanas
        word.DisplayAlerts = False

        count = 0

        if not os.path.exists(plantillas_dir):
            print(f"Carpeta no encontrada: {plantillas_dir}")
            return

        print("Iniciando escaneo de documentos...")

        for filename in os.listdir(plantillas_dir):
            if filename.endswith(".docx") and not filename.startswith("~$"):
                filepath = os.path.join(plantillas_dir, filename)
                try:
                    doc = word.Documents.Open(filepath)

                    # 2. Recorremos todas las partes del documento (Cabeceras, Cuerpo, Pies de página)
                    for old_text, new_text in replacements.items():
                        for story_range in doc.StoryRanges:
                            # Reemplaza en el bloque actual
                            story_range.Find.Execute(
                                FindText=old_text,
                                MatchCase=False,
                                MatchWholeWord=False,
                                ReplaceWith=new_text,
                                Replace=2 # 2 = wdReplaceAll
                            )

                            # Si el bloque tiene formas o cuadros de texto enlazados, iteramos en ellos
                            next_range = story_range.NextStoryRange
                            while next_range is not None:
                                next_range.Find.Execute(
                                    FindText=old_text,
                                    ReplaceWith=new_text,
                                    Replace=2
                                )
                                next_range = next_range.NextStoryRange

                    doc.Save()
                    doc.Close(SaveChanges=True)
                    print(f"[OK] Editado automáticamente: {filename}")
                    count += 1
                except Exception as ex:
                    print(f"[ERROR] Falló en {filename}: {str(ex)}")

        word.Quit()
        print(f"\n¡Éxito! Se inyectaron variables en {count} documentos.")
    except Exception as e:
        print("Error crítico al interactuar con Word:", e)
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    mass_replace()
