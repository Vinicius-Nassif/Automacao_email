from pathlib import Path

from docx2pdf import convert     #pip install docx2pdf
import pandas as pd 	#pip install pandas openpyxl
from docxtpl import DocxTemplate 	#pip install docxtpl

base_dir = Path(__file__).parent
word_template_path = base_dir / "vert_contato.docx"
excel_path = base_dir / "contacts-list.xlsx"
output_dir = base_dir / "OUTPUT"


## Criar pasta output para os documentos word
# Create output folder for the word documents
output_dir.mkdir(exist_ok=True)


## Converter planilha do Excel em Pandas Dataframe
# Convert Excel sheet into pandas dataframe
df = pd.read_excel(excel_path, sheet_name="Sheet1")


##  Exibir somente data YYYY-MM-DD (sem o horário)
# Keep only date part YYYY-MM-DD (not the time)
df["DATA_ENTREGA"] = pd.to_datetime(df["DATA_ENTREGA"]).dt.date


## Alterando o formato da data de YYYY-MM-DD para DD-MM-YYYY
# Changing format date YYYY-MM-DD to DD-MM-YYYY
df["DATA_ENTREGA"] = pd.to_datetime(df["DATA_ENTREGA"]).dt.strftime('%d/%m/%Y')


## Interação entre o Excel e a produção dos documentos Word
# Interate over each row in df and render word document 
for record in df.to_dict(orient="records"):
	doc = DocxTemplate(word_template_path)
	doc.render(record)
	output_path = output_dir / f"{record['NOME_DESTINATARIO']} - vert_contato.docx"
	doc.save(output_path)

## Conversão de .docx para .pdf
# Covert .docx into .pdf
	convert( output_path, f"{output_dir}/{record['NOME_DESTINATARIO']}.pdf")
