import os
import re
import pdfplumber
import pandas as pd

# === CONFIGURAÇÕES ===
PASTA_PROJETO = os.getcwd()
PASTA_DOWNLOAD = os.path.join(PASTA_PROJETO, "das_baixados")

def extrair_linha_digitavel(caminho_pdf):
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            # O DAS geralmente tem 1 página. Pegamos o texto todo.
            texto = pdf.pages[0].extract_text()
            
            # Limpeza básica: remover quebras de linha e espaços extras
            texto_limpo = texto.replace("\n", "").replace(" ", "")
            
            # REGEX: Procura sequências que começam com 85 (Impostos) e têm muitos dígitos
            # O DAS tem 4 blocos, mas no PDF o texto pode vir junto.
            match = re.search(r'85\d{40,}', texto_limpo)
            
            if match:
                return match.group(0)
            else:
                return "Código não encontrado"
    except Exception as e:
        return f"Erro: {e}"

# Lista para guardar os dados
dados_finais = []

# Varre a pasta de downloads
print("Lendo arquivos PDF...")
arquivos = [f for f in os.listdir(PASTA_DOWNLOAD) if f.endswith(".pdf")]

for arquivo in arquivos:
    cnpj = arquivo.replace(".pdf", "") # O nome do arquivo é o CNPJ
    caminho_completo = os.path.join(PASTA_DOWNLOAD, arquivo)
    
    print(f"Lendo: {arquivo}...")
    linha_digitavel = extrair_linha_digitavel(caminho_completo)
    
    dados_finais.append({
        "CNPJ": cnpj,
        "CODIGO_BARRAS": linha_digitavel,
        "DATA_PAGAMENTO": "20/02/2025", # Pode automatizar isso depois
        "VALOR": "checar no banco" # Opcional: extrair valor com regex também
    })

# Salva o Excel Final
df_resultado = pd.DataFrame(dados_finais)
df_resultado.to_excel("Lote_Pagamento_Final.xlsx", index=False)

print("\n------------------------------------------------")
print("SUCESSO! Arquivo 'Lote_Pagamento_Final.xlsx' gerado.")
print("Abra este arquivo e copie os códigos para o banco.")