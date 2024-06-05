# Projeto: Automação de Extração de Dados de Notebooks da Magazine Luiza

## Descrição
Este projeto automatiza a tarefa de extração de dados de notebooks listados no site da Magazine Luiza. O robô realiza a pesquisa de produtos, coleta informações específicas e organiza os dados em um arquivo Excel. Além disso, o robô envia um e-mail com o relatório gerado.

## Funcionalidades
1. Verificar se o site da Magazine Luiza carregou corretamente.
2. Realizar até três tentativas de carregamento do site. Se o erro persistir, gerar um log com a mensagem "Site fora do ar".
3. Pesquisar por "notebooks" no campo de pesquisa.
4. Coletar o nome do produto, a quantidade de avaliações e a URL de cada notebook listado na página de resultados.
5. Filtrar e organizar os dados coletados em um arquivo Excel:
    - Remover produtos que não possuem avaliações.
    - Classificar os produtos em duas categorias: "Melhores" (mais de 100 avaliações) e "Piores" (menos de 100 avaliações).
6. Salvar os dados organizados em um arquivo Excel nomeado "Notebooks.xlsx" dentro de uma pasta chamada "Output".
7. Enviar um e-mail com o relatório gerado em anexo.

## Estrutura do Projeto
```
projeto-raiz/
│
├── Config.py
├── Objects/
│   ├── Obj_EmailSender.py
│   └── Obj_WebAutomation.py
│
├── Output/
│   ├── Notebooks.xlsx
│   └── erro_carregamento.log
│
├── main.py
│
└── README.md
```

## Requisitos
- Python 3.x
- Bibliotecas Python:
  - `selenium`
  - `pandas`
  - `beautifulsoup4`
  - `smtplib`
  - `email`
- WebDriver compatível com seu navegador (ChromeDriver para Google Chrome, por exemplo)

## Instalação
1. Clone o repositório para sua máquina local.
2. Instale as dependências necessárias usando pip:
   ```bash
   pip install selenium pandas beautifulsoup4 webdriver_manager openpyxl
   ```
3. Baixe o WebDriver correspondente ao seu navegador e adicione o caminho do WebDriver ao PATH do sistema, ou especifique o caminho diretamente no código.

## Como Executar
1. Configure o arquivo `Config.py` com suas informações de e-mail:
   ```python
   user_email = 'seu-nome'
   password = 'sua-senha'
   email = 'seu-email@gmail.com'
   ```
2. Execute o script principal:
   ```bash
   python main.py
   ```

## Estrutura do Script

### Verificação de Carregamento do Site
```python
def verify_website(webdriver: WebDriver, url: str, tries: int = 3) -> bool:
    for tr in range(tries):
        try:
            webdriver.get(url)
            return True
        except Exception:
            print(f"Tentativa {tr + 1} falhou, tentando novamente...")

    return False
```

### Pesquisa de Produtos
```python
def search_product(driver: Driver, webdriver: WebDriver, product: str) -> None:
    input_field = driver.find_by_element(webdriver, '//input[@id="input-search"]', wait=5)
    input_field.send_keys(product)
    driver.click_by_element(webdriver, '//div[@data-testid="input-container"]//*[name()="svg"]')
    driver.find_by_element(webdriver, '//span[@data-testid="main-title"]', wait=10)
```

### Extração de Dados
```python
def extrair_dados(driver: Driver, webdriver: WebDriver, max_pages=100) -> list:
    produtos = []
    pag_atual = 1
    while pag_atual < max_pages:
        html = webdriver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        product_list = soup.find('div', {'data-testid': 'product-list'})
        products: list[BeautifulSoup] = product_list.find_all('li')
        for product in products:
            nome = product.find('h2', {'data-testid': 'product-title'}).text
            url = product.find('a', {'data-testid': 'product-card-container'}).get('href')
            url = f'https://www.magazineluiza.com.br{url}'
            aval = product.find('span', {'data-testid': 'review'})
            qtd_aval = (0 if not aval else aval.text.split()[1].removeprefix('(').removesuffix(')'))
            produtos.append([nome, qtd_aval, url])
        print(f"Pagina {pag_atual} concluída")
        try:
            driver.find_by_element(webdriver, '//button[@type="next" and @disabled]')
            break
        except Exception:
            pag_atual += 1
            driver.click_by_element(webdriver, '//button[@type="next"]', wait=5)
            while f'page={pag_atual}' not in webdriver.current_url:
                sleep(2)
            sleep(3)
    return produtos
```

### Criação do DataFrame
```python
def criar_dataframe(produtos) -> pd.DataFrame:
    df = pd.DataFrame(produtos, columns=['PRODUTO', 'QTD_AVAL', 'URL'])
    df = df[df['QTD_AVAL'].astype(int) > 0]
    return df
```

### Salvamento no Excel
```python
def salvar_excel(df: pd.DataFrame) -> None:
    df_melhores = df[df['QTD_AVAL'].astype(int) >= 100]
    df_piores = df[df['QTD_AVAL'].astype(int) < 100]
    with pd.ExcelWriter('Output/Notebooks.xlsx') as writer:
        df_melhores.to_excel(writer, sheet_name='Melhores', index=False)
        df_piores.to_excel(writer, sheet_name='Piores', index=False)
```

### Envio de E-mail
```python
def enviar_email(to_email: list, subject: list, body: str) -> None:
    email = Email(subject)
    email.sender = user_email
    email.destination = to_email
    email.body = body.replace('\n', '<br>')
    email.smtp_host = 'smtp.gmail.com'
    email.smtp_port = 587
    email.__password__ = password
    email.add_atachment('Output/Notebooks.xlsx', 'Relatório.xlsx', 'vnd.ms-excel')
    email.create_email()
    email.send_email(confirm_send_message=True)
```

### Função Principal
```python
def main() -> None:
    url = 'https://www.magazineluiza.com.br'
    to_list = [user_email]
    email_body = """
        Olá, aqui está o seu relatório dos notebooks
        extraídos da Magazine Luiza.

        Atenciosamente,

        Robô.
    """
    driver = Driver()
    webdriver = driver.new_driver(no_window=True)
    if not verify_website(webdriver, url):
        webdriver.quit()
        with open('Output/erro_carregamento.log', 'w') as log_file:
            log_file.write('Site fora do ar')
    search_product(driver, webdriver, 'notebooks')
    produtos = extrair_dados(driver, webdriver)
    df = criar_dataframe(produtos)
    salvar_excel(df)
    enviar_email(to_list, 'Relatório de Notebooks', email_body)
```

### Considerações Finais
Este projeto é um exemplo prático de automação de tarefas web utilizando Selenium e Python. Sinta-se à vontade para personalizar e expandir este projeto conforme suas necessidades.

### Contribuição
Contribuições são bem-vindas! Por favor, abra um problema ou envie uma solicitação de pull request para melhorias e correções.

### Licença
Este projeto está licenciado sob a licença MIT. Consulte o arquivo LICENSE para obter mais detalhes.

### Contato
Para dúvidas e sugestões, entre em contato pelo e-mail [nycolaspimentel12@gmail.com](mailto:nycolaspimentel12@gmail.com)
