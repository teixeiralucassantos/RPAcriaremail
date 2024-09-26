# Envio de E-mail com Outlook usando Python

Este script em Python permite a criação e o envio de e-mails utilizando a biblioteca `win32com.client`, que interage com o Microsoft Outlook. O exemplo incluído no código é um e-mail de feliz aniversário, mas a estrutura pode ser facilmente adaptada para qualquer tipo de mensagem.

## Funcionalidades

- Verifica se o Outlook está em execução.
- Cria um novo e-mail com destinatário, assunto e corpo em HTML.
- Permite anexar um arquivo, como uma assinatura.
- Exibe o e-mail para revisão antes de enviar (opcional).
- Salva o e-mail como rascunho ou o envia diretamente.

## Como Funciona

1. **Verificação do Outlook**: A função `verificar_outlook_aberto` tenta acessar a aplicação Outlook. Se o Outlook não estiver em execução, uma mensagem de erro é exibida, e o script não continua.

2. **Criação do E-mail**: A função `criar_email` é responsável por compor o e-mail. Os parâmetros incluem:
   - `destinatario`: o endereço de e-mail para o qual a mensagem será enviada.
   - `assunto`: o assunto do e-mail.
   - `corpo_html`: o corpo da mensagem em formato HTML, permitindo formatação rica.
   - `anexo`: um caminho opcional para um arquivo que será anexado ao e-mail.

3. **Uso de HTML**: O corpo do e-mail é estruturado em HTML, permitindo uma apresentação visual mais atraente. Você pode personalizar o conteúdo como desejar.

4. **Anexos**: Se um caminho para um arquivo anexo for fornecido e o arquivo existir, ele será anexado ao e-mail.

5. **Exibição e Envio**: O e-mail é exibido para o usuário antes de ser enviado. Isso é útil para revisão. O e-mail pode ser salvo como rascunho ou enviado diretamente, dependendo das suas necessidades.

## Exemplo de Uso

No exemplo presente no código, um e-mail é criado para celebrar o aniversário de uma pessoa chamada Maria. O corpo do e-mail contém mensagens de carinho e felicitações. 

Para personalizar para outros contextos, basta modificar os parâmetros passados à função `criar_email`.

```python
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
