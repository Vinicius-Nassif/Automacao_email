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

![1](https://user-images.githubusercontent.com/111388699/200914888-7a447366-2db1-4902-9801-34d6097212ac.png)

Primeiramente, ocorreu a importação da *pathlib*, que é uma biblioteca para minipular caminhos de sistemas de arquivos de maneira independente, seja qual for o sistema operacional.

A biblioteca *docxtpl* foi utilizada para preencher os dados desejados dentro do template .docx. 



### 3.2 Código 

#### def __init__(self, doc_inicial, documento_gerado):

O código teve início com a denominação de uma classe chamada *SingleWordGenerator()* e sua primeira função foi a __init__():, como podemos ver a seguir:

![2](https://user-images.githubusercontent.com/111388699/200914922-2de94040-c64d-49ed-b53f-103121e90de5.png)

O método *“__init__():*” é um método especial para que o Python execute automaticamente sempre que criarmos uma nova instância baseada na classe *SingleWordGenerator():* e foi definido para ter três argumentos: o *self, doc_inicial* e *documento_gerado*.

Logo abaixo, iniciam-se os objetos *self.word_template_path*, *self.base_dir*, *self.doc* e *self.context* como vazios com objetivo de receberem seus seus valores no desenvolvimento do código. 



#### def identificacao(self):

Essa função estabelece diretórios dos arquivos e da template. 

![3](https://user-images.githubusercontent.com/111388699/200914954-126620b0-ddc9-4abe-b4d4-cf5f0342114b.png)

A variável *self.base_dir* recebe o método *Path()* para estabelecer o diretório como o mesmo do arquivo *self.doc_inicial*.

A variável *self.word_template_path* recebeu o *self.base_dir / self.doc_inicial* para determinar o seu diretório como o mesmo. Já a variável *self.doc* recebeu o método  *DocxTemplate()* com o parâmetro *self.word_template_path* para determinar qual seria o template a ser trabalhada.

Na sequência, foram fornecidas as informações pela planilha Excel que preecheriam as lacunas identificadas no template à variável *self.context*. Vejamos o template:

![4](https://user-images.githubusercontent.com/111388699/200914994-a45974ef-0e8c-408e-9d15-6a4bc7f19195.png)



#### def render_save(self):

Essa função é responsável por renderizar as informações e salvar o documento com as informações inseridas no template.

![5](https://user-images.githubusercontent.com/111388699/200915026-1eff36c5-863b-4700-9477-f1f1444028dd.png)

A variável *self.doc* recebeu o método *render()* com o parâmetro *self.context* para inserir as informações informadas no novo documento, que receberá o seu valor na variável *self.documento_gerado* e será salvo no mesmo diretório. 



#### if __name__=='__main__':

Podemos usar um bloco **if __name__ ==** "**__main__**" para permitir ou evitar que partes do código sejam executadas ao importar os módulos. Quando o interpretador do *Python* lê um arquivo, a variável **__name__** é definida como **__main__** se o módulo que está sendo executado, ou como o nome do módulo se ele for importado.

![6](https://user-images.githubusercontent.com/111388699/200915072-8645a76b-4701-4432-b0a2-12f97668f444.png)

Assim, tem o objetivo de orquestrar a execução sequencial de todas as fases do single-word-generator, como um índice representativo e intuitivo, com o chamado de todos os métodos já explicados e trazendo também mensagens de êxito nas conclusões de cada etapa pelo método *print()*. Fica assim o resultado após o preenchimento da template:

![7](https://user-images.githubusercontent.com/111388699/200915106-856defce-ec1a-48a0-a1d4-4ea663fde828.png)



## 4. bulk-word-generator

Esse é o módulo responsável por inteirar no template e produzir uma série documentos .docx com as informações extraídas da planilha Excel.

### 4.1 Bibliotecas 

![8](https://user-images.githubusercontent.com/111388699/200915165-7025ea7b-2c3e-416c-933f-bb6a7066c625.png)

Primeiramente, ocorreu a importação da *pathlib*, que é uma biblioteca para minipular caminhos de sistemas de arquivos de maneira independente, seja qual for o sistema operacional.

A biblioteca *docx2pdf* instrumentalizou a conversão dos arquivos .docx em .pdf após a renderização dos documentos com as respectivas informações da planilha Excel.  Foi importada a biblioteca *Pandas* como a sigla“pd”, com o objetivo de manipulação e análise de dados.

A biblioteca *docxtpl* foi utilizada para preencher os dados desejados dentro do template .docx. 



### 4.2 Código

#### def __init__(self, documento_inicial, planilha, diretorio):

O código teve início com a denominação de uma classe chamada *BulkWordGenerator()* e sua primeira função foi a __init__():, como podemos ver a seguir:

![9](https://user-images.githubusercontent.com/111388699/200915233-580a2e09-a1a6-43df-becb-4081b9160900.png)

O método *“__init__():*” é um método especial para que o Python execute automaticamente sempre que criarmos uma nova instância baseada na classe *BulkWordGenerator():* e foi definido para ter quatro argumentos: o self, *documento_inicial* e *planilha* e *diretorio*.

Logo abaixo, iniciam-se os objetos *self.word_template_path*, *self.excel_path*, *self.output_dir* *self.base_dir* e *self.df*  como vazios com objetivo de receberem seus seus valores no desenvolvimento do código. 



#### def identificacao(self):

Essa função estabelece diretórios dos arquivos, da template e da planilha Excel.

![10](https://user-images.githubusercontent.com/111388699/200915279-7f97cd6a-981f-4a2a-b8d6-5bea88b538f7.png)

A variável self.base_dir recebe o método *Path()* para estabelecer o diretório como o mesmo do arquivo *self.doc_inicial*.

A variável *self.word_template_path* recebeu o *self.base_dir / self.documento_inicial* para determinar o seu diretório como o mesmo. Já a variável *self.excel_path* recebeu como diretório o *self.base_dir* e a *self.planilha*.

Por último, o *self.output_dir* recebeu a variável *self.base_dir* e o *self.diretorio*, que específicam o diretório inicial e o de destino dos arquivos a serem renderizados. 



#### def manipulacao(self):

É o método responsável em criar a pasta de destino dos arquivos e converter a planilha Excel em *dataframe*.

![11](https://user-images.githubusercontent.com/111388699/200915331-03f13d9f-d615-48a1-a58f-010a65be1b80.png)

A variável *self.output_dir* recebeu o método *mkdir*() para criar o diretório de destino. A variável *self.df* recebeu o *pandas* como pd com o método *read_excel()*, especificando o argumento com a variável self.*excel_path* e o nome da planilha como "Sheet1". 

Em caráter ilustrativo, vejamos parte da planilha que contém as informações:

![plan](https://user-images.githubusercontent.com/111388699/200915420-e42e2812-bd15-4067-8aeb-be0b9d0abccd.png)



#### def data(self):

Essa função ficou responsável por exibir a data nos documentos renderizados.

![12](https://user-images.githubusercontent.com/111388699/200915477-600dff97-1e00-44db-b16a-fa5d39bac509.png)

Foi estabelecido para a variável *self.df* que exiba no campo do template "DATA_ENTREGA" receba a data no formato brasileiro (DD/MM/YY). 



#### def interacao(self):

Essa função ficou responsável por interar entre o Excel e a produção dos .docx.

![13](https://user-images.githubusercontent.com/111388699/200915524-a8a94471-895e-424b-be5f-c46e671aa0ed.png)

Se inicia com um laço de repetição for para que em *record* em *self.df* receba o *dataframe*. Foi identificado o template na variável doc recebendo o método DocxTemplate() e sua renderização na linha 43. 

A variável *output_path* recebeu o *self.output_dir* com o formato de nome do arquivo sendo no *NOME_DESTINATARIO - vert_coontatos.docx*, sendo esse documento salvo no *output_path*. Na sequênia, exibe uma mensagem de conclusão.

Na linha 49 ocorre a conversão dos arquivos do output_path com o formato do NOME_DESTINARIO.pdf. Na sequência exibe a mensagem de conclusão.

![15](https://user-images.githubusercontent.com/111388699/200915636-09ef28af-1353-4eb2-934d-3ee8a560fb41.png)



#### def __name__=="__main__":

Tem o objetivo de orquestrar a execução sequencial de todas as fases do single-word-generator, como um índice representativo e intuitivo, com o chamado de todos os métodos já explicados e trazendo também mensagens de êxito nas conclusões de cada etapa pelo método *print()*. 

![14](https://user-images.githubusercontent.com/111388699/200915688-ee4f2a91-aa9c-443d-b8ef-cbbceba42c16.png)

A pasta output após a conclusão do procedimento de automação  passou a conter arquivos .docx e .pdf de cada cliente. 
![16](https://user-images.githubusercontent.com/111388699/200915732-c3d5c030-9ce6-4636-baf3-73c66d7a3f1b.png)



## 5. Conclusão

​		Esse projeto foi criado para melhor elucidar o conhecimento trazido pela automação de template, pois no seu desenvolvimento foi necessária a interação de várias bibliotecas do Python, bem como entender conceitos e aplicações da lógica de programação.

​		Percebeu-se também que para solucionar o problema principal foi necessário dividi-lo em pequenas tarefas para o desenvolvimento. 

​		Portanto, essa é uma forma de usar a manipulação de dados a seu favor, seja para uso pessoal ou profissional.
