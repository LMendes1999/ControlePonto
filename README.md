# Controle de ponto em planilha Google Sheets
Esta classe implementa uma aplicação para controle de ponto em uma planilha do Google Sheets. A aplicação é escrita em Java e utiliza a API do Google Drive e do Google Sheets para manipular os dados na planilha. A aplicação tem as seguintes funcionalidades:

- Autenticação do usuário
- Busca pela planilha com nome padrão "DEMONSTRATIVO DAS HORAS CONTRATADAS" no Google Drive
- Busca pela data atual na coluna A da planilha
- Insere um registro de ponto com a data e hora atual na primeira célula vazia das colunas D, E, F, G, H ou I da linha correspondente à data encontrada

A aplicação utiliza as bibliotecas Google Client API, Google OAuth Client e Google OAuth Client Jetty. A classe ControlePonto contém os métodos para autenticação do usuário, autorização para acesso à API do Google Drive e do Google Sheets, busca da planilha com nome padrão e inserção do registro de ponto na planilha. Os métodos são implementados utilizando a API do Google.

A autenticação do usuário é realizada utilizando o fluxo de autorização do OAuth2. A API do Google Drive e do Google Sheets requerem uma credencial válida para serem acessadas. Para obter essa credencial, a aplicação utiliza o fluxo de autorização do OAuth2 para solicitar que o usuário conceda permissão à aplicação para acessar seus dados.

Após a autorização do usuário, a aplicação cria um serviço do Google Drive e um serviço do Google Sheets que podem ser usados para acessar os dados da planilha. A aplicação busca pela planilha com nome padrão "DEMONSTRATIVO DAS HORAS CONTRATADAS" no Google Drive utilizando a API do Google Drive e retorna o ID da planilha.

A aplicação então utiliza a API do Google Sheets para buscar pela data atual na coluna A da planilha. Se a data atual é encontrada, a aplicação procura pela primeira célula vazia nas colunas D, E, F, G, H ou I da linha correspondente à data encontrada e insere um registro de ponto com a data e hora atual. Se a data atual não é encontrada na planilha, a aplicação exibe uma mensagem de erro.

Os métodos implementados na classe ControlePonto são executados na ordem em que são definidos na classe principal. Se alguma das etapas falhar, a aplicação exibe uma mensagem de erro e interrompe a execução.

> **Importante:**
> Essa aplicação utiliza uma credencial privada para solicitar consentimentos de uso, para compilar o projeto corretamente, é necessário incluir no projeto uma credencial com escopo do Google Drive e do Google Sheets
