# Empregare Scraper – Busca de Vagas com Exportação para Excel

Script em Python que faz web scraping no site **Empregare**, permitindo buscar vagas por **cargo**, **localidade** e **quantidade de rolagens** da página, e gera um **arquivo Excel profissional** com todas as vagas encontradas, incluindo **link clicável** para acessar e se cadastrar.

## Funcionalidades

- Entrada interativa: cargo, localidade e número de rolagens (10 vagas por rolagem, em média).
- Coleta de várias vagas na página de resultados.
- Geração de arquivo `xlsx` com:
  - Cabeçalhos formatados (título em negrito e fundo azul).
  - Coluna final com **hiperlink** para a vaga (“🔗 Abrir Vaga”).
- Organização ideal para análise de oportunidades em concursos/empregos.

## Tecnologias e Bibliotecas Utilizadas

Este projeto foi pensado para demonstrar habilidades com **Python para automação, scraping e manipulação de planilhas**.

### Selenium

- Utilizado para **controlar o navegador**, realizar buscas no site Empregare e rolar a página automaticamente.
- Permite:
  - Preencher campos de busca (cargo, localidade).
  - Clicar em botões e aguardar carregamento dinâmico de vagas.
  - Executar rolagens configuráveis pelo usuário (scroll da página).

### BeautifulSoup (beautifulsoup4)

- Usado para **parsear o HTML retornado** e extrair as informações das vagas.
- Responsável por:
  - Encontrar os elementos do HTML que representam cada vaga.
  - Extrair dados como título, empresa, localização, salário (quando disponível) e link da vaga.

### Pandas

- Utilizado para **organizar os dados em um DataFrame** antes de exportar para Excel.
- Facilita:
  - Estruturar as vagas em formato tabular.
  - Tratar listas/dicionários de vagas de forma limpa.
  - Exportar rapidamente com `DataFrame.to_excel()`.

### Openpyxl

- Usado para **pós-processar e formatar** o arquivo Excel gerado.
- Responsável por:
  - Formatar o **cabeçalho** (fontes em negrito, cor de fundo azul, texto branco).
  - Criar uma **tabela do Excel** com estilo “TableStyleLight9” (tabela clara azul).
  - Ajustar largura das colunas manualmente para melhor leitura.
  - Criar **hiperlinks** na última coluna (G) com o texto “🔗 Abrir Vaga”, apontando para o link real da vaga.

## Estrutura do Projeto

Exemplo de estrutura de arquivos:

```text
empregare-scraper/
├─ empregare_scraper.py
├─ requirements.txt
└─ README.md
```

- `empregare_scraper.py`: script principal com toda a lógica (input do usuário, Selenium, BeautifulSoup, Pandas e formatação com openpyxl).
- `requirements.txt`: dependências do projeto, incluindo versões usadas.

## Requisitos

Conteúdo típico do `requirements.txt`:

```txt
beautifulsoup4==4.14.3
openpyxl==3.1.5
pandas==2.3.3
selenium==4.39.0
```

Instalação:

```bash
pip install -r requirements.txt
```

## Como Usar

1. Clone o repositório:  
   ```bash
   git clone https://github.com/Drewthyson/WebScraping.git
   cd empregare-scraper
   ```

2. Crie e ative um ambiente virtual (opcional, mas recomendado).

3. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

4. Execute o script:
   ```bash
   python empregare_scraper.py
   ```

5. Informe:
   - Cargo desejado (ex.: “Desenvolvedor Python”).
   - Localidade (ex.: “Brasília”).
   - Quantidade de rolagens na página.

6. Ao final, um arquivo Excel (ex.: `Vagas_Empregare.xlsx`) será gerado com:
   - Todas as vagas encontradas.  
   - Cabeçalho formatado.  
   - Tabela estilizada.  
   - Links clicáveis para cada vaga.

## Observações

- Projeto desenvolvido com foco em **portfólio** e demonstração de:
  - Automatização de navegador com Selenium.
  - Web scraping com BeautifulSoup.
  - Manipulação de dados com Pandas.
  - Geração e formatação avançada de Excel com openpyxl.

***