from pathlib import Path
from docx2pdf import convert     #pip install docx2pdf
import pandas as pd 	#pip install pandas 
from docxtpl import DocxTemplate 	#pip install docxtpl

class BulkWordGenerator():

	def __init__(self, documento_inicial, planilha, diretorio):
		# Alocando argumentos
		self.documento_inicial = documento_inicial
		self.planilha = planilha 
		self.diretorio = diretorio

		# Inicializando objetos
		self.word_template_path = None
		self.excel_path = None
		self.output_dir = None
		self.base_dir = None
		self.df = None

	def identicacao(self):
		self.base_dir = Path(__file__).parent
		self.word_template_path = self.base_dir / self.documento_inicial
		self.excel_path = self.base_dir / self.planilha
		self.output_dir = self.base_dir / self.diretorio

	def manipulacao(self):
		# Criar pasta output para os documentos word
		self.output_dir.mkdir(exist_ok=True)
		# Converter planilha do Excel em Pandas Dataframe
		self.df = pd.read_excel(self.excel_path, sheet_name="Sheet1")

	def data(self):
		# Exibir somente data YYYY-MM-DD (sem o horário)
		self.df["DATA_ENTREGA"] = pd.to_datetime(self.df["DATA_ENTREGA"]).dt.date
		# Alterando o formato da data de YYYY-MM-DD para DD-MM-YYYY
		self.df["DATA_ENTREGA"] = pd.to_datetime(self.df["DATA_ENTREGA"]).dt.strftime('%d/%m/%Y')

	def interacao(self):
		# Interação entre o Excel e a produção dos documentos Word
		for record in self.df.to_dict(orient="records"):
			doc = DocxTemplate(self.word_template_path)
			doc.render(record)
			output_path = self.output_dir / f"{record['NOME_DESTINATARIO']} - vert_contato.docx"
			doc.save(output_path)
			print('.docx armazenado com sucesso!')

			# Converter .docx em .pdf
			convert(output_path, f"{self.output_dir}/{record['NOME_DESTINATARIO']}.pdf")
			print('.docx convertido para .pdf')

vert_to_excel = BulkWordGenerator(documento_inicial='vert_contato.docx', planilha='contacts-list.xlsx', diretorio='OUTPUT')

if __name__=='__main__':
	vert_to_excel = BulkWordGenerator(documento_inicial='vert_contato.docx', planilha='contacts-list.xlsx', diretorio='OUTPUT')
	# Execução sequencial de todas as fases da classe BulkWordGenerator:
	# 1. Identifica o diretório base do .docx e .xlsx, a template e o output 
	vert_to_excel.identicacao()
	print('Identificação executada com sucesso')
	# 2. Cria a pasta output e converte a planilha do Excel em dataframe
	vert_to_excel.manipulacao()
	print('Manipulação executada com sucesso')
	# 3. Executa a formatação da data
	vert_to_excel.data()
	print('Datação realizada com sucesso')
	# 4. Interação entre Excel e os documentos .docx. Depois converte em .pdf
	vert_to_excel.interacao()
	print('Interação realizada com sucesso')
	print('BulkWordGenerator concluído!')