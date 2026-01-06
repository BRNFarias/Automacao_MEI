import os
import re
import pandas as pd
import pdfplumber

# === CONFIGURAÇÕES ===
PASTA_DAS = os.path.join(os.getcwd(), "das_baixados")
ARQUIVO_SAIDA = "lista_pagamentos_itau.xlsx"

def limpar_texto(texto):
    """Remove tudo que não for número"""
    return re.sub(r'[^0-9]', '', texto)

def processar_pdf(caminho_arquivo):
    """Lê o PDF e extrai CNPJ, Valor e Linha Digitável"""
    nome_arquivo = os.path.basename(caminho_arquivo)
    dados = {
        "Arquivo": nome_arquivo,
        "CNPJ": "Não encontrado",
        "Valor": "0,00",
        "Codigo_Barras": "Não encontrado",
        "Status": "Erro"
    }

    try:
        with pdfplumber.open(caminho_arquivo) as pdf:
            texto_completo = ""
            for pagina in pdf.pages:
                texto_completo += pagina.extract_text() or ""
            
            # --- 1. Extração do CNPJ ---
            # Procura padrão XX.XXX.XXX/XXXX-XX
            match_cnpj = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', texto_completo)
            if match_cnpj:
                dados["CNPJ"] = match_cnpj.group(0)

            # --- 2. Extração da Linha Digitável (Ouro) ---
            # DAS (Guias de Arrecadação) tem 48 dígitos e começam com 8
            # O texto no PDF geralmente vem quebrado com espaços ou hifens
            # Estratégia: Limpar tudo e buscar sequência de 48 números começando com 8
            
            texto_limpo = limpar_texto(texto_completo)
            
            # Regex: Procura sequência de 48 dígitos que começa com 8
            match_barras = re.search(r'8\d{47}', texto_limpo)
            
            if match_barras:
                dados["Codigo_Barras"] = match_barras.group(0)
                dados["Status"] = "OK"
            else:
                dados["Status"] = "Sem código de barras"

            # --- 3. Extração de Valor (Para conferência) ---
            # Tenta pegar o valor bruto. Geralmente aparece como "Total R$ 100,00"
            match_valor = re.search(r'Total\s+R\$\s+([\d\.,]+)', texto_completo)
            if match_valor:
                dados["Valor"] = match_valor.group(1)

    except Exception as e:
        dados["Status"] = f"Erro leitura: {str(e)}"
        
    return dados

# === EXECUÇÃO PRINCIPAL ===
if __name__ == "__main__":
    print(f"--- EXTRATOR PARA O ITAÚ ---")
    print(f"Lendo pasta: {PASTA_DAS}")

    if not os.path.exists(PASTA_DAS):
        print(f"ERRO: Pasta '{PASTA_DAS}' não existe. Rode o robô baixador primeiro.")
        exit()

    arquivos_pdf = [f for f in os.listdir(PASTA_DAS) if f.lower().endswith(".pdf")]
    
    if not arquivos_pdf:
        print("Nenhum PDF encontrado na pasta.")
        exit()

    lista_dados = []

    print(f"Processando {len(arquivos_pdf)} arquivos...")

    for arquivo in arquivos_pdf:
        caminho_completo = os.path.join(PASTA_DAS, arquivo)
        resultado = processar_pdf(caminho_completo)
        
        lista_dados.append(resultado)
        
        # Feedback visual simples
        simbolo = "✅" if resultado["Status"] == "OK" else "❌"
        print(f"{simbolo} {arquivo} -> {resultado['Codigo_Barras']}")

    # --- SALVAR EXCEL ---
    if lista_dados:
        df = pd.DataFrame(lista_dados)
        
        # Reordenar colunas para facilitar cópia
        colunas = ["CNPJ", "Valor", "Codigo_Barras", "Arquivo", "Status"]
        df = df[colunas]
        
        df.to_excel(ARQUIVO_SAIDA, index=False)
        
        print("\n" + "="*40)
        print(f"RELATÓRIO GERADO COM SUCESSO!")
        print(f"Arquivo: {ARQUIVO_SAIDA}")
        print("="*40)
        print("COMO PAGAR NO ITAÚ:")
        print("1. Abra o Excel gerado.")
        print("2. Copie a coluna 'Codigo_Barras'.")
        print("3. No Itaú Empresas, vá em Pagamentos > Boletos/Títulos > Digitação.")
        print("4. Cole a lista. O banco vai reconhecer os valores automaticamente.")