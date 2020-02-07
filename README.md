# Tutorial Power Pivot | Habilitar + Exemplo Prático


Este tutorial visa demonstrar de forma simplificada a utilização do **Power Pivot**, assim sendo, será utilizado um exemplo bastante simples. Para um tutorial aprofundado, clique [aqui](https://support.office.com/pt-br/article/come%C3%A7ar-a-usar-o-power-pivot-no-microsoft-excel-fdfcf944-7876-424a-8437-1a6c1043a80b).

### Antes de começar, vamos a uma introdução sobre o que é o Power Pivot.
O **Power Pivot** é um recurso do **Microsoft Excel**, está disponível como um suplemento. Para habilita-lo, basta seguir alguns passos simples que serão demonstrados a seguir. Como o **Power Pivot** é possível criar *modelos de dados*, *estabelecer relações* e *criar cálculos*. Você pode trabalhar com grandes conjuntos de dados, estabelecer relações extensas e criar cálculos complexos ou simples. Tudo isso em um ambiente de alto desempenho e com a experiência clássica do **Excel**. O **Power Pivot** é uma das três ferramentas de análise disponíveis no **Excel**, além dela temos o **Power Query** e o **Power View**. Mas neste tutorial será obordado apenas o **Power Pivot**.

## Habilitando o Power Pivot
01. Com o *Excel* aberto, clique em **Arquivo**;

![1](https://user-images.githubusercontent.com/38192454/73904340-89015f80-487a-11ea-8f02-1fa0cc076081.PNG)
 
 
02. Clique em **Opções**;

![2](https://user-images.githubusercontent.com/38192454/73904358-93bbf480-487a-11ea-80a9-24b086ff73c1.PNG)


03. Clique em **Suplementos**;

![3](https://user-images.githubusercontent.com/38192454/73904367-99b1d580-487a-11ea-99ea-243258b2cdd3.PNG)


04. Em **Gerenciar**, selecione **Suplementos COM** e clique em **Ir...**;

![4](https://user-images.githubusercontent.com/38192454/73904376-9e768980-487a-11ea-9ca3-64d9c7463a59.PNG)


05. Marque a opção **Microsoft Power Pivot for Excel** e clique em **Ok**;

![5](https://user-images.githubusercontent.com/38192454/73904382-a3d3d400-487a-11ea-9606-b57d00bcc659.PNG)


06. Observe que a *Guia* do **Power Pivot** foi habilitada.

![6](https://user-images.githubusercontent.com/38192454/73904393-aa624b80-487a-11ea-9094-ef0e938b6587.PNG)


## Exemplo Prático

**Como exemplo prático, queremos saber quantos alunos de Sistemas de Informação (SI) são da Turma de 2017.**


01. Clique na *Guia* do **Power Pivot**, conforme imagem anterior;

02. Clique em **Gerenciar**, a *Janela* do **Power Pivot** será aberta;

![7](https://user-images.githubusercontent.com/38192454/73904398-ae8e6900-487a-11ea-9f73-2ec593ba9de9.PNG)


03. Para adicionar **Dados**, vá em **Obter Dados Externos** e clique em **De Outras Fontes**;

![8](https://user-images.githubusercontent.com/38192454/73904402-b51ce080-487a-11ea-8002-a7c5e9e39416.PNG)


Para o exemplo aqui proposto, foi utilizado como fonte de dados um arquivo do **Excel** chamado **SAD.xlsx** que está disponível neste repositório.

04. Desça a *Barra de Rolagem*, selecione **Arquivo do Excel** e clique em **Avançar >**;

![9](https://user-images.githubusercontent.com/38192454/73904407-b8b06780-487a-11ea-8212-d4223bf1043f.PNG)

05. Em **Caminho do arquivo do Excel:**, clique em **Procurar** é selecione o arquivo **SAD.xlsx**. Marque à opção **Usar primeira linha como cabeçalho da coluna**, e clique em **Avançar >**;

![10](https://user-images.githubusercontent.com/38192454/74057762-5d3dc100-49c3-11ea-896f-c0f7db231c11.PNG)

06. Clique em **Visualizar e Filtrar**;

![11](https://user-images.githubusercontent.com/38192454/74057767-5f078480-49c3-11ea-8ed5-f72f7fc903da.PNG)

07. Aqui é possível fazer alguns filtros. Para filtrar *colunas*, basta desmarcar o *checkbox*. Para o exemplo que utilizaremos não precisamos das colunas **Situacao** e **Prioridade**. Desmarque elas;

![12](https://user-images.githubusercontent.com/38192454/74057770-6038b180-49c3-11ea-86d8-48971546342b.PNG)

08. Para filtrar *linhas*, clique no botão localizado do lado direito do nome da coluna, e faça o passo seguinte;

![13](https://user-images.githubusercontent.com/38192454/74057772-62027500-49c3-11ea-9e8f-3a3e6e199c69.PNG)

09. Para o exemplo proposto, queremos apenas os alunos de SI, então desmarque o *checkbox* referente à **CIENCIA DA COMPUTACAO**, e clique em **Ok**;

![14](https://user-images.githubusercontent.com/38192454/74057775-629b0b80-49c3-11ea-84d7-9c91edca0dd9.PNG)

10. Clique em **OK**;

![15](https://user-images.githubusercontent.com/38192454/74057776-6333a200-49c3-11ea-9ce2-d72877feb51f.PNG)

11. Clique em **Concluir**;

![16](https://user-images.githubusercontent.com/38192454/74057779-64fd6580-49c3-11ea-9b8f-d059328cf686.PNG)

12. Após a importação ser concluida, clique em **Fechar**;

![17](https://user-images.githubusercontent.com/38192454/74057782-662e9280-49c3-11ea-9c7a-c1ee8315eed7.PNG)

13. Clique em um dos campos embaixo da tabela (conforme à seta) e clique na **Barra de Fórmulas**. Insira à seguinte *string* sem aspas: "**T:=COUNTROWS(FILTER('SAD'; [Ano]=2017))**", e aperte **Enter** no teclado;

![18](https://user-images.githubusercontent.com/38192454/74057784-66c72900-49c3-11ea-9d17-a115151302f3.PNG)

14. O total de alunos de SI da turma de 2017 é 4.

![19](https://user-images.githubusercontent.com/38192454/74057788-6890ec80-49c3-11ea-8f4a-d4717bf3a22b.PNG)


## Entendendo a *String*

A *string* esta escrita em **DAX** (*Data Analysis Expressions*), que é uma linguagem de fórmula. Para não estender muito vou apenas explicar à *string*, caso queira obter mais informações sobre **DAX**, clique [aqui](https://docs.microsoft.com/pt-br/dax/dax-overview).
- **T**: O nome da medida. Fórmulas para medidas podem incluir o nome da medida, seguido por dois pontos, seguido da fórmula de cálculo.
- **=**: O operador de sinal de igual indica o início da fórmula (o segundo igual, indica comparação).
- **( )**: Parênteses envolvem um ou mais argumentos. Todas as funções exigem pelo menos um argumento. Um argumento passa um valor para uma função.
- [COUNTROWS](https://docs.microsoft.com/pt-br/dax/countrows-function-dax): Função que conta o número de linhas na tabela especificada ou em uma tabela definida por uma expressão.
- [FILTER](https://docs.microsoft.com/pt-br/dax/filter-function-dax): Função que filtra conforme o argumento passado.
- **' '**: Dentro de aspas simples, é colocado o nome da tabela referenciada.
- **[ ]**: Dentro de colchetes, é colocado o nome da coluna referenciada.


## Especificação do Notebook utilizado para o presente Tutorial
|Sistema Operacional|Memória RAM|Processador|
|----------------------|----------------------|----------------------|
| Windows 10 Home |  4GB DDR4  |Inter(R) Core(TM) i5-7200U CPU @ 2.5GHz|
