import win32com.client as client
import os

def verificar_outlook_aberto():
    """Verifica se o Outlook está em execução."""
    try:
        client.Dispatch("Outlook.Application")
        return True
    except Exception as e:
        print(f"Erro ao acessar o Outlook: {e}")
        return False

def criar_email(destinatario, assunto, corpo_html, anexo=None):
    """Cria e envia um e-mail com o Outlook."""
    if not verificar_outlook_aberto():
        print("O Outlook não está em execução. Feche e reabra o Outlook e tente novamente.")
        return

    # Criando a instância do Outlook dentro da função
    outlook = client.Dispatch("Outlook.Application")
    novo_email = outlook.CreateItem(0)
    novo_email.To = destinatario
    novo_email.Subject = assunto
    novo_email.HTMLBody = corpo_html

    # Anexando a assinatura se existir
    if anexo and os.path.exists(anexo):
        novo_email.Attachments.Add(anexo)

    # Exibindo o email antes de enviar (opcional)
    novo_email.Display()  # Pode ser comentado se quiser enviar diretamente

    # Salvar ou enviar
    novo_email.Save()  # Salva como rascunho
    # novo_email.Send()  # Descomente esta linha para enviar diretamente


# Caminho para a assinatura
assinatura_caminho = os.path.join(os.getcwd(), 'Assinatura.jpeg')

# Dados do e-mail
destinatario = "maria.silva@example.com"  # Novo destinatário
assunto = "Celebração do seu Aniversário!"
corpo_html = """
    <h2>Feliz Aniversário, Maria!</h2>
    <p>Hoje é o seu dia especial! Que seja repleto de amor e alegria.</p>
    <p>Que você tenha um dia maravilhoso ao lado das pessoas que ama.</p>
    <p>Desejamos a você muita saúde, sucesso e felicidade!</p>
    <p>Contamos com você para celebrarmos juntos!</p>
    <br>
    <p>Com carinho,</p>
    <p>Sua equipe de sempre.</p>
"""

# Criando e enviando o e-mail
criar_email(destinatario, assunto, corpo_html, assinatura_caminho)
