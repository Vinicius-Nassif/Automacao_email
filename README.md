# Projeto Automação de Template para E-mail

## 1. Introdução

A automação é o uso de tecnologia para automatizar ações e processos, reduzindo os trabalhos manuais para aumentar a eficiência na comunicação com seus contatos desejados. 

Nesse projeto, o *Python* foi adotado como linguagem de programação para instrumentalizar o serviço de um freelancer responsável que deveria, a partir de um template, automatizar a criação de documentos .pdf individualizados para serem enviados por e-mail em momento futuro tendo como referência uma planilha Excel contendo como informações:

- Nome do destinatário;
- Endereço;
- CEP;
- Data da entrega;
- Número da OC; e
- Produto.

O trabalho teve como paradigma uma empresa de tecnologia e venda de equipamentos eletrônicos. Essa empresa faz entregas à domicílio e por muitas vezes acaba desencontrando com o cliente. Quando isso ocorre, o produto volta para a loja e espera o contato do cliente para reagendar a entrega. Assim, é produzido o documento que será enviado por e-mail padrão contendo as informações personalizadas. 

O intuito dessa atividade irá melhor expor a automação de processos, divisão de módulos, variáveis, laço de repetição e manipulação de arquivos, conforme demonstrado a seguir.



## 2. Módulos 

Módulo é um arquivo contendo funções em *Python* para serem usados em outros programas da mesma linguagem. 

Para instrumentalizar essa atividade de maneira organizada, foi realizada em dois módulos de código *Python*:

- single-word-generator; e
- bulk-word-generator.

Cada um será responsável por diferentes atividades do processo e serão abordados no decorrer desse trabalho. 



## 3. single-word-generator

Esse é o módulo responsável por inteirar no template e produzir somente um documento .docx com as informações desejadas.

### 3.1 Bibliotecas 

O projeto teve início com a importação das bibliotecas e funções utilizadas na elaboração do código em *Python*:

![Imagem 1](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\1.png)

Primeiramente, ocorreu a importação da *pathlib*, que é uma biblioteca para minipular caminhos de sistemas de arquivos de maneira independente, seja qual for o sistema operacional.

A biblioteca *docxtpl* foi utilizada para preencher os dados desejados dentro do template .docx. 



### 3.2 Código 

#### def __init__(self, doc_inicial, documento_gerado):

O código teve início com a denominação de uma classe chamada *SingleWordGenerator()* e sua primeira função foi a __init__():, como podemos ver a seguir:

![Imagem 2](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\2.png)

O método *“__init__():*” é um método especial para que o Python execute automaticamente sempre que criarmos uma nova instância baseada na classe *SingleWordGenerator():* e foi definido para ter três argumentos: o *self, doc_inicial* e *documento_gerado*.

Logo abaixo, iniciam-se os objetos *self.word_template_path*, *self.base_dir*, *self.doc* e *self.context* como vazios com objetivo de receberem seus seus valores no desenvolvimento do código. 



#### def identificacao(self):

Essa função estabelece diretórios dos arquivos e da template. 

![Imagem 3](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\3.png)

A variável *self.base_dir* recebe o método *Path()* para estabelecer o diretório como o mesmo do arquivo *self.doc_inicial*.

A variável *self.word_template_path* recebeu o *self.base_dir / self.doc_inicial* para determinar o seu diretório como o mesmo. Já a variável *self.doc* recebeu o método  *DocxTemplate()* com o parâmetro *self.word_template_path* para determinar qual seria o template a ser trabalhada.

Na sequência, foram fornecidas as informações pela planilha Excel que preecheriam as lacunas identificadas no template à variável *self.context*. Vejamos o template:

![Imagem 4](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\4.png)



#### def render_save(self):

Essa função é responsável por renderizar as informações e salvar o documento com as informações inseridas no template.

![Imagem 5](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\5.png)

A variável *self.doc* recebeu o método *render()* com o parâmetro *self.context* para inserir as informações informadas no novo documento, que receberá o seu valor na variável *self.documento_gerado* e será salvo no mesmo diretório. 



#### if __name__=='__main__':

Podemos usar um bloco **if __name__ ==** "**__main__**" para permitir ou evitar que partes do código sejam executadas ao importar os módulos. Quando o interpretador do *Python* lê um arquivo, a variável **__name__** é definida como **__main__** se o módulo que está sendo executado, ou como o nome do módulo se ele for importado.

![Imagem 6](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\6.png)

Assim, tem o objetivo de orquestrar a execução sequencial de todas as fases do single-word-generator, como um índice representativo e intuitivo, com o chamado de todos os métodos já explicados e trazendo também mensagens de êxito nas conclusões de cada etapa pelo método *print()*. Fica assim o resultado após o preenchimento da template:

![Imagem 7](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\7.png)



## 4. bulk-word-generator

Esse é o módulo responsável por inteirar no template e produzir uma série documentos .docx com as informações extraídas da planilha Excel.

### 4.1 Bibliotecas 

![Imagem 8](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\8.png)

Primeiramente, ocorreu a importação da *pathlib*, que é uma biblioteca para minipular caminhos de sistemas de arquivos de maneira independente, seja qual for o sistema operacional.

A biblioteca *docx2pdf* instrumentalizou a conversão dos arquivos .docx em .pdf após a renderização dos documentos com as respectivas informações da planilha Excel.  Foi importada a biblioteca *Pandas* como a sigla“pd”, com o objetivo de manipulação e análise de dados.

A biblioteca *docxtpl* foi utilizada para preencher os dados desejados dentro do template .docx. 



### 4.2 Código

#### def __init__(self, documento_inicial, planilha, diretorio):

O código teve início com a denominação de uma classe chamada *BulkWordGenerator()* e sua primeira função foi a __init__():, como podemos ver a seguir:

![Imagem 9](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\9.png)

O método *“__init__():*” é um método especial para que o Python execute automaticamente sempre que criarmos uma nova instância baseada na classe *BulkWordGenerator():* e foi definido para ter quatro argumentos: o self, *documento_inicial* e *planilha* e *diretorio*.

Logo abaixo, iniciam-se os objetos *self.word_template_path*, *self.excel_path*, *self.output_dir* *self.base_dir* e *self.df*  como vazios com objetivo de receberem seus seus valores no desenvolvimento do código. 



#### def identificacao(self):

Essa função estabelece diretórios dos arquivos, da template e da planilha Excel.

![Imagem 10](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\10.png)

A variável self.base_dir recebe o método *Path()* para estabelecer o diretório como o mesmo do arquivo *self.doc_inicial*.

A variável *self.word_template_path* recebeu o *self.base_dir / self.documento_inicial* para determinar o seu diretório como o mesmo. Já a variável *self.excel_path* recebeu como diretório o *self.base_dir* e a *self.planilha*.

Por último, o *self.output_dir* recebeu a variável *self.base_dir* e o *self.diretorio*, que específicam o diretório inicial e o de destino dos arquivos a serem renderizados. 



#### def manipulacao(self):

É o método responsável em criar a pasta de destino dos arquivos e converter a planilha Excel em *dataframe*.

![Imagem 11](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\11.png)

A variável *self.output_dir* recebeu o método *mkdir*() para criar o diretório de destino. A variável *self.df* recebeu o *pandas* como pd com o método *read_excel()*, especificando o argumento com a variável self.*excel_path* e o nome da planilha como "Sheet1". 

Em caráter ilustrativo, vejamos parte da planilha que contém as informações:

![Imagem Plan](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\plan.png)

#### def data(self):

Essa função ficou responsável por exibir a data nos documentos renderizados.

![Imagem 12](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\12.png)

Foi estabelecido para a variável *self.df* que exiba no campo do template "DATA_ENTREGA" receba a data no formato brasileiro (DD/MM/YY). 



#### def interacao(self):

Essa função ficou responsável por interar entre o Excel e a produção dos .docx.

![Imagem 13](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\13.png)

Se inicia com um laço de repetição for para que em *record* em *self.df* receba o *dataframe*. Foi identificado o template na variável doc recebendo o método DocxTemplate() e sua renderização na linha 43. 

A variável *output_path* recebeu o *self.output_dir* com o formato de nome do arquivo sendo no *NOME_DESTINATARIO - vert_coontatos.docx*, sendo esse documento salvo no *output_path*. Na sequênia, exibe uma mensagem de conclusão.

Na linha 49 ocorre a conversão dos arquivos do output_path com o formato do NOME_DESTINARIO.pdf. Na sequência exibe a mensagem de conclusão.

![Imagem 15](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\15.png)



#### def __name__=="__main__":

Tem o objetivo de orquestrar a execução sequencial de todas as fases do single-word-generator, como um índice representativo e intuitivo, com o chamado de todos os métodos já explicados e trazendo também mensagens de êxito nas conclusões de cada etapa pelo método *print()*. 

![Imagem 14](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\14.png)

A pasta output após a conclusão do procedimento de automação  passou a conter arquivos .docx e .pdf de cada cliente. 

![Imagem 16](C:\Users\ccece\Documents\Projetos AD\Artigos\Automacao_email\16.png)



## 5. Conclusão

​		Esse projeto foi criado para melhor elucidar o conhecimento trazido pela automação de template, pois no seu desenvolvimento foi necessária a interação de várias bibliotecas do Python, bem como entender conceitos e aplicações da lógica de programação.

​		Percebeu-se também que para solucionar o problema principal foi necessário dividi-lo em pequenas tarefas para o desenvolvimento. 

​		Portanto, essa é uma forma de usar a manipulação de dados a seu favor, seja para uso pessoal ou profissional.