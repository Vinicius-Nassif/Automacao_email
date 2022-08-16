import datetime
from pathlib import Path


from docxtpl import DocxTemplate # pip install docxtpl

## Criando variáveis para path absoluto do repositório
base_dir = Path(__file__).parent
word_template_path = base_dir / "vert_contato.docx"

doc = DocxTemplate(word_template_path)
context = {
	"NOME_DESTINATARIO": "André Almeida fontes",
	"ENDEREÇO": "Rua Alameda dos santos, número 5",
	"CEP": "65232-232",
	"DATA_ENTREGA": "23/05/2021",
	"NUMERO_OC": "1200323134",
	"PRODUTO": "TECLADO GAMER XL",
}
doc.render(context)
doc.save(base_dir / "vert_contato_generated.docx")