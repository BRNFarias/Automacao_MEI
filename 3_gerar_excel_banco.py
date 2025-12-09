import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime

# === CONFIGURAÇÕES ===
PASTA_DAS = "das_baixados"
ARQUIVO_SAIDA = "Lista_Para_Pagamento.xlsx"
DATA_PAGAMENTO = "20/02/2025" # Ajuste para a data que você vai pagar

def extrair_dados_pdf(caminho_pdf):
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            # Remove quebras de linha para facilitar a busca
            texto_limpo = texto.replace("\n", "").replace(" ", "")
            
            # 1. Busca Linha Digitável (Código de Barras - começa com 85 e tem aprox 48 digitos)
            match_barras = re.search(r'85\d{40,}', texto_limpo)
            codigo = match_barras.group(0) if match_barras else "ERRO_LEITURA"
            
            # 2. Busca Valor (Opcional - tenta achar o padrão R$ XX,XX)
            # No DAS o valor total costuma vir após "Valor Total"
            # Essa busca é genérica, o mais importante é o código de barras
            valor = "0,00"
            match_valor = re.search(r'TotalR\$(\d+,\d{2})', texto_limpo)
            if match_valor:
                valor = match_valor.group(1)
            
            return codigo, valor
    except Exception as e:
        return "ERRO_ARQUIVO", str(e)

# === EXECUÇÃO ===
print(f"--- Iniciando processamento da pasta: {PASTA_DAS} ---")

if not os.path.exists(PASTA_DAS):
    print(f"ERRO: A pasta '{PASTA_DAS}' não existe. Rode o script 1 primeiro.")
    exit()

arquivos = [f for f in os.listdir(PASTA_DAS) if f.endswith(".pdf")]

if not arquivos:
    print("Nenhum PDF encontrado! Você baixou as guias?")
    exit()

lista_final = []

for arquivo in arquivos:
    print(f"Lendo: {arquivo}...")
    caminho = os.path.join(PASTA_DAS, arquivo)
    
    # O nome do arquivo é o CNPJ (ex: 12345678000199.pdf)
    cnpj = arquivo.replace(".pdf", "")
    
    codigo_barras, valor = extrair_dados_pdf(caminho)
    
    status = "OK"
    if codigo_barras == "ERRO_LEITURA":
        status = "VERIFICAR MANUALMENTE"
        
    lista_final.append({
        "Identificador (CNPJ)": cnpj,
        "Código de Barras": codigo_barras,
        "Valor Estimado": valor,
        "Data Pagamento": DATA_PAGAMENTO,
        "Status": status
    })

# Gera o Excel
df = pd.DataFrame(lista_final)
df.to_excel(ARQUIVO_SAIDA, index=False)

print("\n" + "="*40)
print(f"SUCESSO! Arquivo gerado: {ARQUIVO_SAIDA}")
print("Passos Finais:")
print("1. Abra o arquivo 'Lista_Para_Pagamento.xlsx'.")
print("2. Copie a coluna 'Código de Barras'.")
print("3. Cole na área de 'Pagamento em Lote' do seu Banco (Inter/Cora/Nubank).")