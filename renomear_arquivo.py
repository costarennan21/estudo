import os
import shutil
from datetime import datetime, timedelta

# Configurações
diretorio = r"C:\Users\RennanCosta\OneDrive - Diretório Padrão\Relatório SLA Rede D'Or"
nome_base = "Acompanhamento Validação Time REDE D'Or.xlsx.xlsm"  # NOME CORRETO!
data_ontem = (datetime.now() - timedelta(days=0)).strftime('%d.%m')  # Formato DD.MM
novo_nome = f"Acompanhamento Validação Time REDE D'Or {data_ontem}.xlsx.xlsm"  # Padrão existente

# Caminhos completos
caminho_original = os.path.join(diretorio, nome_base)
caminho_novo = os.path.join(diretorio, novo_nome)

# Verifica se o arquivo original existe
if not os.path.exists(caminho_original):
    print(f"❌ Arquivo base não encontrado: {nome_base}")
    print("\n📂 Arquivos existentes:")
    for arq in os.listdir(diretorio):
        print(f" - {arq}")
    exit()

# Se o arquivo de destino já existir, adiciona um sufixo (_2, _3, etc.)
contador = 2
while os.path.exists(caminho_novo):
    novo_nome = f"Acompanhamento Validação Time REDE D'Or {data_ontem}_{contador}.xlsx.xlsm"
    caminho_novo = os.path.join(diretorio, novo_nome)
    contador += 1

# Cria a cópia (mantém o original)
try:
    shutil.copy2(caminho_original, caminho_novo)
    print(f"✅ Cópia criada com sucesso: {novo_nome}")
    print(f"   Original preservado: {nome_base}")
except PermissionError:
    print("❌ Erro: O arquivo está aberto no Excel. Feche-o e tente novamente.")
except Exception as e:
    print(f"❌ Erro inesperado: {e}")