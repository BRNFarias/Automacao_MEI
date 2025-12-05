import os
import re
import pdfplumber
import pandas as pd

# Configurações
PASTA_DAS = "das_baixados"
ARQUIVO_SAIDA = "Lista_Para_Pagamento.xlsx"

def extrair_dados_do_das(caminho_pdf):
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            
            # Limpa o texto para facilitar a busca (remove espaços e quebras)
            texto_limpo = texto.replace("\n", "").replace(" ", "")
            
            # 1. Busca Linha Digitável (Código de Barras)
            # Começa com 85... e tem aprox 48 digitos
            match_barras = re.search(r'85\d{40,}', texto_limpo)
            codigo = match_barras.group(0) if match_barras else "NÃO ENCONTRADO"
            
            # 2. Busca Valor (Opcional, busca padrão R$ XX,XX)
            # Regex simples para pegar valores monetários pode ser complexo, 
            # mas o código de barras já contém o valor.
            
            return codigo
    except Exception as e:
        return f"Erro: {e}"

# Lista para o Excel
dados = []

print(f"Lendo arquivos na pasta '{PASTA_DAS}'...")

if not os.path.exists(PASTA_DAS):
    print("Pasta não encontrada!")
    exit()

arquivos = [f for f in os.listdir(PASTA_DAS) if f.endswith(".pdf")]

for arquivo in arquivos:
    print(f"Lendo: {arquivo}...")
    caminho = os.path.join(PASTA_DAS, arquivo)
    
    # O nome do arquivo é o CNPJ (se o passo 1 funcionou bem)
    cnpj = arquivo.replace(".pdf", "")
    
    codigo_barras = extrair_dados_do_das(caminho)
    
    dados.append({
        "CNPJ": cnpj,
        "CODIGO_BARRAS": codigo_barras,
        "DATA_VENCIMENTO": "20/02/2025", # Pode automatizar ou deixar fixo
        "STATUS": "OK" if codigo_barras != "NÃO ENCONTRADO" else "VERIFICAR PDF"
    })

# Gera o Excel
df = pd.DataFrame(dados)
df.to_excel(ARQUIVO_SAIDA, index=False)

print("\n" + "="*40)
print(f"SUCESSO! Arquivo '{ARQUIVO_SAIDA}' gerado.")
print("Agora é só subir esse arquivo no seu Banco (Inter, Cora, etc) em 'Pagamento em Lote'.")