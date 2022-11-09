from pathlib import Path
from docxtpl import DocxTemplate # pip install docxtpl

class SingleWordGenerator():

	def __init__(self, doc_inicial, documento_gerado):
		# Alocando argumentos
		self.doc_inicial = doc_inicial
		self.documento_gerado = documento_gerado

		#Inicializando objetos
		self.word_template_path = None
		self.base_dir = None
		self.doc = None
		self.context = None

	def identificacao(self):
		# Estabelecendo como mesmo diretório do self.doc_inicial
		self.base_dir = Path(__file__).parent
		# Determinando o diretório e a template 
		self.word_template_path = self.base_dir / self.doc_inicial
		self.doc = DocxTemplate(self.word_template_path)
		# Fornecendo informações que preencherão as lacunas identificadas na template
		self.context = {
					"NOME_DESTINATARIO": "André Almeida fontes",
					"ENDEREÇO": "Rua Alameda dos santos, número 5",
					"CEP": "65232-232",
					"DATA_ENTREGA": "23/05/2021",
					"NUMERO_OC": "1200323134",
					"PRODUTO": "TECLADO GAMER XL",
					   }

	def render_save(self):
		# Renderizando as informações no template
		self.doc.render(self.context)
		# Salvando o documento com as informações inseridas no template
		self.doc.save(self.base_dir / self.documento_gerado)

if __name__=='__main__':
	vert_documentacao = SingleWordGenerator(
				doc_inicial='vert_contato.docx', 
				documento_gerado='vert_contato_generated.docx')
	# Execução sequencial de todas as fases da classe SingleWordGenerator:
	# 1. Identicação do documento inicial, do diretório e lacunas a serem preenchidas.
	vert_documentacao.identificacao()
	print('Identificação executada com sucesso')
	# 2. Renderização e salvamento do arquivo .docx gerado
	vert_documentacao.render_save()
	print('Documento renderizado e armazenado com sucesso')
	print('Vert Documentação concluída!')