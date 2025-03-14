!pip install python-docx
!apt-get update
!apt-get install libreoffice # Removed extra spaces before this line
import pandas as pd
import platform
import subprocess
import shutil
from docx import Document
from docx.shared import Cm

# ... (rest of your code remains the same)

import pandas as pd
import platform
import subprocess
import shutil
from docx import Document
from docx.shared import Cm

def extract_data_from_excels(excel_files, distrito, provincia, departamento):
    data_by_file = {}
    for excel_file in excel_files:
        try:
            df = pd.read_excel(excel_file)
            required_cols = {'Distrito', 'Provincia', 'Departamento'}
            if not required_cols.issubset(df.columns):
                print(f"Advertencia: {excel_file} no contiene todas las columnas requeridas {required_cols}")
                continue

            df_filtered = df[(df['Distrito'].str.lower() == distrito.lower()) &
                             (df['Provincia'].str.lower() == provincia.lower()) &
                             (df['Departamento'].str.lower() == departamento.lower())]
            data_by_file[excel_file] = df_filtered
        except Exception as e:
            print(f"Error procesando {excel_file}: {e}")
    return data_by_file

def replace_text_in_document(doc, replacements):
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in replacements.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, value)

def insert_table(doc, data, title):
    doc.add_paragraph(f"\n\n**{title}:**")
    table = doc.add_table(rows=1, cols=len(data.columns) + 1)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Nro."
    hdr_cells[0].width = Cm(0.83)

    for i, column_name in enumerate(data.columns):
        hdr_cells[i + 1].text = column_name

    for idx, (_, row) in enumerate(data.iterrows(), start=1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[0].width = Cm(0.83)
        for i, value in enumerate(row):
            row_cells[i + 1].text = str(value)

def convert_to_pdf(input_docx, output_pdf):
    system_os = platform.system()

    try:
        if system_os == "Windows":
            from docx2pdf import convert
            convert(input_docx)
        elif system_os in ["Linux", "Darwin"]:
            libreoffice_path = shutil.which("libreoffice") or shutil.which("soffice")
            if libreoffice_path:
                subprocess.run([libreoffice_path, "--headless", "--convert-to", "pdf", input_docx])
            else:
                print("Error: LibreOffice no está instalado o no se encuentra en la ruta del sistema.")
    except Exception as e:
        print(f"Error en la conversión a PDF: {e}")

def generate_final_docx(output_docx, output_pdf, plantilla_docx, distrito, provincia, departamento, tables_by_file):
    doc = Document(plantilla_docx)

    selected_columns = ['PobrezaMonetaria', 'CCPP', 'Poblacion', 'Hogares', 'Viviendas', 'UnidadTerritorial']
    data_dict = {col: 'N/A' for col in selected_columns}

    if "datos1.xlsx" in tables_by_file and not tables_by_file["datos1.xlsx"].empty:
        data_dict.update(tables_by_file["datos1.xlsx"][selected_columns].iloc[0].to_dict())

    focalizado, seleccionado, mejorando, social = "NO", "NO", "NO", "NO"
    total_centros_poblados = "0"
    centros_poblados_df, infraestructura_social_df = pd.DataFrame(), pd.DataFrame()

    if "datos2.xlsx" in tables_by_file:
        table_cp = tables_by_file["datos2.xlsx"]
        if not table_cp.empty and all(col in table_cp.columns for col in ['Provincia', 'Distrito', 'CentroPoblado']):
            table_cp_filtered = table_cp[table_cp['Distrito'].str.lower() == distrito.lower()]
            focalizado = "SI" if not table_cp_filtered.empty else "NO"
            total_centros_poblados = str(len(table_cp_filtered))
            centros_poblados_df = table_cp_filtered[['Provincia', 'Distrito', 'CentroPoblado']]

    if "datos3.xlsx" in tables_by_file and 'Distrito' in tables_by_file["datos3.xlsx"]:
        seleccionado = "SI" if distrito.lower() in tables_by_file["datos3.xlsx"]['Distrito'].str.lower().values else "NO"

    if "datos4.xlsx" in tables_by_file and 'Distrito' in tables_by_file["datos4.xlsx"]:
        mejorando = "SI" if distrito.lower() in tables_by_file["datos4.xlsx"]['Distrito'].str.lower().values else "NO"

    if "datos5.xlsx" in tables_by_file:
        table_soc = tables_by_file["datos5.xlsx"]
        if not table_soc.empty and all(col in table_soc.columns for col in ['Provincia', 'Distrito', 'CentroPoblado']):
            social = "SI"
            infraestructura_social_df = table_soc[['Provincia', 'Distrito', 'CentroPoblado']]

    replacements = {
        "{Distrito}": distrito,
        "{Provincia}": provincia,
        "{Departamento}": departamento,
        "{PobrezaMonetaria}": str(data_dict.get("PobrezaMonetaria", "N/A")),
        "{CCPP}": str(data_dict.get("CCPP", "N/A")),
        "{Poblacion}": str(data_dict.get("Poblacion", "N/A")),
        "{Hogares}": str(data_dict.get("Hogares", "N/A")),
        "{Viviendas}": str(data_dict.get("Viviendas", "N/A")),
        "{UnidadTerritorial}": str(data_dict.get("UnidadTerritorial", "N/A")),
        "{Focalizado}": focalizado,
        "{Seleccionado}": seleccionado,
        "{Mejorando}": mejorando,
        "{Social}": social,
        "{TotalCentrosPoblados}": total_centros_poblados
    }
    replace_text_in_document(doc, replacements)

    if focalizado == "SI" and not centros_poblados_df.empty:
        insert_table(doc, centros_poblados_df, "Lista de Población Objetivo 2025")

    if social == "SI" and not infraestructura_social_df.empty:
        insert_table(doc, infraestructura_social_df, "Lista de Infraestructura Social")

    doc.save(output_docx)
    convert_to_pdf(output_docx, output_pdf)

distrito = "SANTILLANA"
provincia = "HUANTA"
departamento = "AYACUCHO"
excel_files = ["datos1.xlsx", "datos2.xlsx", "datos3.xlsx", "datos4.xlsx", "datos5.xlsx"]
plantilla_docx = "plantilla.docx"
output_docx = f"AyudaMemoria_{distrito}.docx"
output_pdf = f"AyudaMemoria_{distrito}.pdf"

tables_by_file = extract_data_from_excels(excel_files, distrito, provincia, departamento)
generate_final_docx(output_docx, output_pdf, plantilla_docx, distrito, provincia, departamento, tables_by_file)

print(f"Informe generado: {output_docx} y convertido a {output_pdf}")
