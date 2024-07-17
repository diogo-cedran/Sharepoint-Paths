# Sharepoint-Paths

Automação em Python para tratar a limitação de caracteres no caminho dos arquivos do Sharepoint

## Introdução

Conforme estipulado na [documentação oficial da Microsoft](https://docs.microsoft.com/), a plataforma SharePoint impõe uma restrição de 400 caracteres no caminho dos arquivos. Esta restrição é aplicada à combinação do caminho da pasta e do nome do arquivo, após a decodificação. Além disso, o Windows impõe uma limitação de 260 caracteres para o caminho dos arquivos. Esses limites podem causar problemas significativos para os usuários, especialmente quando arquivos e pastas são sincronizados localmente, resultando em caminhos de arquivos que excedem esses limites.

Para solucionar essa "dor do cliente", desenvolvi um código em Python que automatiza a tarefa de percorrer todas as bibliotecas de documentos em um site do SharePoint, listar todos os caminhos de pastas, subpastas e arquivos, e contabilizar a quantidade de caracteres de cada caminho. Isso é útil para identificar caminhos de arquivos que podem ultrapassar os limites permitidos pelo Windows ou pelo SharePoint Online, ajudando na manutenção e organização dos arquivos e pastas dentro do SharePoint. É importante notar que o código funciona apenas em um tenant sem autenticação multifator (MFA).

## Funcionamento do Código

O script autentica-se no SharePoint usando as credenciais fornecidas e percorre todas as bibliotecas de documentos do site raiz. Para cada biblioteca, ele processa recursivamente todas as pastas e arquivos, armazenando os caminhos completos e a quantidade de caracteres de cada um. O script utiliza uma abordagem de checkpoint para salvar o progresso e permitir a retomada do processamento em caso de interrupções, garantindo que nenhum arquivo ou pasta seja processado mais de uma vez.

## Passo a Passo para Utilização do Código

### Instalação da Biblioteca

A biblioteca `Office365-REST-Python-Client` é necessária para interagir com o SharePoint. Essa biblioteca permite autenticar e interagir com o SharePoint Online utilizando a API REST do Office 365.

Para instalar a biblioteca, abra o terminal e execute:

```sh
pip install Office365-REST-Python-Client
```

### Criar o Ambiente Virtual

Criar um ambiente virtual é uma boa prática para isolar as dependências do projeto. Um ambiente virtual garante que as bibliotecas necessárias para o projeto não entrem em conflito com outras bibliotecas instaladas globalmente no sistema. Isso ajuda a manter um ambiente de desenvolvimento limpo e controlado.

No terminal, navegue até o diretório onde deseja criar o projeto e execute:

```sh
python -m venv venv
```

Ative o ambiente virtual:

No Windows:

```sh
.\venv\Scripts\activate
```

No macOS/Linux:

```sh
source venv/bin/activate
```

### Configurar Credenciais e URL

Abra o arquivo `sharepoint_path_length.py` e configure as variáveis `url`, `username` e `password` com as suas credenciais do SharePoint.

### Executar o Script

No terminal, com o ambiente virtual ativado e estando no diretório onde o script está localizado, execute:

```sh
python sharepoint_path_length.py
```

O script irá autenticar-se no SharePoint e começar a percorrer todas as bibliotecas de documentos, processando todas as pastas e arquivos.

## O que o Script Vai Gerar

### Autenticação

O script primeiro autentica-se no SharePoint usando as credenciais fornecidas.

### Percorrendo Bibliotecas

O script busca todas as bibliotecas de documentos no site raiz do SharePoint.

### Processamento Recursivo

Para cada biblioteca, o script percorre recursivamente todas as pastas e arquivos, registrando os caminhos completos e a quantidade de caracteres.

### Salvamento de Resultados

Os resultados são salvos em um arquivo CSV chamado `sharepoint_paths.csv`, contendo duas colunas: "Número de Caracteres" e "Caminho da Pasta/Arquivo".

Um arquivo de checkpoint (`checkpoint.txt`) é utilizado para salvar o progresso e permitir a retomada do processamento em caso de interrupção.

Ao final, o arquivo `sharepoint_paths.csv` conterá uma listagem detalhada de todos os caminhos de pastas e arquivos do SharePoint, juntamente com a contagem de caracteres de cada caminho, ajudando a identificar possíveis problemas de comprimento de caminho.
