import io
import smtplib
import warnings

from os.path import join
from PIL import Image
from email.message import EmailMessage
from colorama import init

warnings.filterwarnings('ignore')
init()


class Email:
    def __init__(self, title: str) -> None:
        """Cria um objeto de um e-mail que poderá ser enviado futuramente

        Args:
            title (str): Insira o título do e-mail
        """
        self.subject: str = title
        self.destination: list = []
        self.cc: list = []
        self.bcc: list = []
        self.body: str = None
        self.__email__: EmailMessage = None

        # Configurações do servidor SMTP
        self.smtp_host: str = 'insira seu serviço SMTP'
        self.smtp_port: int = 587
        self.sender: str = 'insira seu e-mail'
        self.__password__: str = 'insira sua senha'

        self.bcc.append(self.sender)
        self.__attachments__: list = []

    def create_email(self) -> EmailMessage:
        if self.sender in self.destination:
            self.destination.remove(self.sender)

        if self.sender in self.cc:
            self.cc.remove(self.sender)

        # Criação da mensagem multipart
        result = EmailMessage()
        result['From'] = self.sender
        result['Subject'] = self.subject
        result['To'] = ", ".join(self.destination)
        result['Cc'] = ", ".join(self.cc)
        result['Bcc'] = ", ".join(self.cc)

        result.set_content(self.body, subtype='html')
        self.__email__ = result

    def add_atachment(self,
                      attachment: str,
                      attachment_name: str,
                      type_file: str,
                      main_type: str = 'application') -> None:

        self.__attachments__.append(
            {
                'path': attachment,
                'name': attachment_name,
                'main_type': main_type,
                'sub_type': type_file
            }
        )

    def add_image(self,
                  image_path: str,
                  image_filename: str,
                  image_cid: str) -> None:

        with open(join(image_path, image_filename), 'rb') as image:
            img_file = image.read()
            img = Image.open(io.BytesIO(img_file))
            image_type = img.format.lower()

        self.__email__.add_related(
            img_file, 'image', image_type, cid=image_cid)

    def send_email(self, confirm_send_message=True):
        self.__get_attachments__()

        with smtplib.SMTP(self.smtp_host, self.smtp_port) as server:
            try:
                server.starttls()
                server.login(self.sender, self.__password__)

                if confirm_send_message:  # Confirma antes de enviar o e-mail
                    res = input(
                        "\033[1;33mdeseja enviar o email? [s/n]: \033[0m")

                    if res.upper() == 'S':
                        server.sendmail(from_addr=self.sender,
                                        to_addrs=(
                                            self.destination +
                                            self.cc +
                                            self.bcc),
                                        msg=self.__email__.as_string())

                        print('\033[1;32mO email foi '
                              'enviado com sucesso!\033[0m')

                else:  # Envia o e-mail diretamente sem confirmação
                    server.sendmail(from_addr=self.sender,
                                    to_addrs=self.destination,
                                    msg=self.__email__.as_string())
                    print('\033[1;32mO email foi enviado com sucesso!\033[0m')

            # Verifica se o e-mail e a senha estão corretos
            except smtplib.SMTPAuthenticationError:
                print("\033[1;31mO e-mail/senha estão incorretos. \033[0m")

    def __get_attachments__(self) -> None:
        for attach in self.__attachments__:
            with open(attach["path"], 'rb') as attachment:
                attachment_data = attachment.read()
                self.__email__.add_attachment(attachment_data,
                                              maintype=attach['main_type'],
                                              subtype=attach['sub_type'],
                                              filename=attach["name"])


if __name__ == '__main__':
    email = Email("Teste", 'Teste de assunto')
    print(email)
