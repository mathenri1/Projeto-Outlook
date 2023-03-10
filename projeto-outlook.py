import imaplib
import schedule

    # Conectar-se ao servidor do Outlook
server = 'outlook.office365.com'
user = 'xxxxxx@xxxxxx.xxx'
password = 'xxxxx'

mail = imaplib.IMAP4_SSL(server)
mail.login(user, password)

def job():
    # Selecionar a pasta "x"
    folder_name = 'xxxxx'
    mail.select(folder_name)

    # Filtrar apenas e-mails não lidos
    _, messages = mail.search(None, 'UNSEEN')

    # Salvar anexos dos e-mails não lidos
    with open('Log de Execução.txt', 'w') as m:
        for message_id in messages[0].split():
            _, message_data = mail.fetch(message_id, '(RFC822)')
            for response_part in message_data:
                if isinstance(response_part, tuple):
                    _, message_body = response_part
                    message = mail.message_from_bytes(message_body)
                    for attachment in message.get_payload():
                        if attachment.get_filename():
                            attachment.SaveAsFile(r"C:\xxxxxx/xxxxx/xxx " + attachment.FileName)
                            m.write(f"Arquivo salvo: {attachment.FileName}\n")

# Executar a tarefa a cada minuto
schedule.every(15).minute.do(job)

while True:
    schedule.run_pending()
