import win32com.client
import os
from datetime import datetime, timedelta
import shutil

def baixar_anexo_outlook():
    # Configurações
    remetente = ""
    nome_anexo_base = ""
    pasta_destino = r"C:\Users\\OneDrive - Diretório Padrão\\"
    
    # Criar pasta de destino se não existir
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)
    
    # Conectar ao Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Pasta Inbox
    
    # Procurar o e-mail mais recente
    mensagens = inbox.Items
    mensagens.Sort("[ReceivedTime]", True)  # Ordenar por data, mais recente primeiro
    
    anexo_baixado = False
    data_limite = datetime.now() - timedelta(days=7)  # Procurar apenas nos últimos 7 dias
    data_atual = datetime.now().strftime("%d-%m-%Y")
    
    for mensagem in mensagens:
        if mensagem.ReceivedTime.replace(tzinfo=None) < data_limite:
            break  # Não verificar e-mails muito antigos
            
        if mensagem.SenderName == remetente or remetente in mensagem.SenderEmailAddress:
            for anexo in mensagem.Attachments:
                if nome_anexo_base in anexo.FileName and anexo.FileName.lower().endswith(('.xls', '.xlsx')):
                    # Definir novo nome com a data
                    extensao = os.path.splitext(anexo.FileName)[1]
                    novo_nome = f"{nome_anexo_base} {data_atual}{extensao}"
                    caminho_completo = os.path.join(pasta_destino, novo_nome)
                    
                    # Verificar se o arquivo já existe e sobrescrever
                    if os.path.exists(caminho_completo):
                        os.remove(caminho_completo)
                    
                    # Salvar o anexo
                    anexo.SaveAsFile(caminho_completo)
                    print(f"Anexo salvo como: {caminho_completo}")
                    anexo_baixado = True
                    break
            
            if anexo_baixado:
                break
    
    if not anexo_baixado:
        print("Nenhum anexo correspondente foi encontrado nos e-mails recentes.")

if __name__ == "__main__":

    baixar_anexo_outlook()
