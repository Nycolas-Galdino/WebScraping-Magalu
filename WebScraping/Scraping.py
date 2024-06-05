import os
import pandas as pd
from time import sleep
from bs4 import BeautifulSoup

from Config import user_email, password

from Objects.Obj_EmailSender import Email
from Objects.Obj_WebAutomation import Driver, WebDriver


def verify_website(webdriver: WebDriver, url: str, tries: int = 3) -> bool:
    for tr in range(tries):
        try:
            webdriver.get(url)
            return True
        except Exception:
            print(f"Tentativa {tr + 1} falhou, tentando novamente...")

    return False


def search_product(driver: Driver, webdriver: WebDriver, product: str) -> None:
    input_field = driver.find_by_element(webdriver, '//input[@id="input-search"]', wait=5)
    input_field.send_keys(product)

    driver.click_by_element(webdriver, '//div[@data-testid="input-container"]//*[name()="svg"]')

    driver.find_by_element(webdriver, '//span[@data-testid="main-title"]',
                           wait=10)


def extrair_dados(driver: Driver, webdriver: WebDriver, max_pages: int = None) -> list:
    produtos = []

    if not max_pages:
        max_pages = 100

    pag_atual = 1
    while pag_atual <= max_pages:
        html = webdriver.page_source
        soup = BeautifulSoup(html, 'html.parser')

        product_list = soup.find('div', {'data-testid': 'product-list'})
        products: list[BeautifulSoup] = product_list.find_all('li')
        for product in products:
            nome = product.find('h2', {'data-testid': 'product-title'}).text
            url = product.find('a', {'data-testid': 'product-card-container'}).get('href')

            url = f'https://www.magazineluiza.com.br{url}'

            aval = product.find('span', {'format': 'score-count'})
            if aval:
                qtd_aval = aval.text.split()[1].removeprefix('(').removesuffix(')')
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


def criar_dataframe(produtos) -> pd.DataFrame:
    df = pd.DataFrame(produtos, columns=['PRODUTO', 'QTD_AVAL', 'URL'])
    df = df[df['QTD_AVAL'].astype(int) > 0]
    return df


def salvar_excel(df: pd.DataFrame) -> None:
    df_melhores = df[df['QTD_AVAL'].astype(int) >= 100]
    df_piores = df[df['QTD_AVAL'].astype(int) < 100]

    with pd.ExcelWriter('Output/Notebooks.xlsx') as writer:
        df_melhores.to_excel(writer, sheet_name='Melhores', index=False)
        df_piores.to_excel(writer, sheet_name='Piores', index=False)


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


def main() -> None:
    url = 'https://www.magazineluiza.com.br'
    to_list = [user_email]
    email_body = """Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza.

        Atenciosamente,
        Robô.
    """

    driver = Driver()
    webdriver = driver.new_driver(no_window=True)

    # Verifica se o site está carregado corretamente
    if not verify_website(webdriver, url):
        print('Site fora do ar, parando o processo...')

        webdriver.quit()
        with open('ErrorLog.log', 'w') as log_file:
            log_file.write('Site fora do ar')
        
        quit()

    # Faz a pesquisa de todos os produtos
    search_product(driver, webdriver, 'notebooks')

    produtos = extrair_dados(driver, webdriver)

    webdriver.quit()

    df = criar_dataframe(produtos)
    salvar_excel(df)
    enviar_email(to_list, 'Relatório Notebooks', email_body)


if __name__ == '__main__':
    os.chdir(os.path.dirname(__file__))
    main()
