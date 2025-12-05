import os
import time
import random
import pandas as pd
from DrissionPage import ChromiumPage, ChromiumOptions

# === CONFIGURAÇÕES ===
PASTA_PROJETO = os.getcwd()
PASTA_DOWNLOAD = os.path.join(PASTA_PROJETO, "das_baixados")
ARQUIVO_CLIENTES = "clientes.xlsx"
ANO_VIGENTE = "2025"

# Cria pasta de download
if not os.path.exists(PASTA_DOWNLOAD):
    os.makedirs(PASTA_DOWNLOAD)

def configurar_pagina():
    co = ChromiumOptions()
    co.set_download_path(PASTA_DOWNLOAD)
    page = ChromiumPage(co)
    return page

def renomear_ultimo_arquivo(cnpj):
    time.sleep(3) 
    arquivos = [os.path.join(PASTA_DOWNLOAD, f) for f in os.listdir(PASTA_DOWNLOAD) if f.endswith(".pdf")]
    
    if not arquivos:
        print("   [!] Nenhum PDF novo encontrado.")
        return
        
    ultimo_arquivo = max(arquivos, key=os.path.getctime)
    novo_nome = os.path.join(PASTA_DOWNLOAD, f"{cnpj}.pdf")
    
    if os.path.exists(novo_nome):
        try: os.remove(novo_nome)
        except: pass
        
    try:
        os.rename(ultimo_arquivo, novo_nome)
        print(f"   [OK] PDF salvo como: {cnpj}.pdf")
    except Exception as e:
        print(f"   [ERRO] Falha ao renomear: {e}")

def espera_aleatoria(min=1, max=3):
    tempo = random.uniform(min, max)
    time.sleep(tempo)

def processar_cliente(page, cnpj):
    print(f"\n--- Processando CNPJ: {cnpj} ---")
    
    page.get("http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Identificacao")
    
    # 1. Login
    if page.ele('#cnpj'):
        page.ele('#cnpj').clear()
        espera_aleatoria(0.5, 1.5)
        page.ele('#cnpj').input(cnpj)
        espera_aleatoria(0.5, 1.0)
        
        if page.ele('#continuar'):
            page.ele('#continuar').click()
    else:
        print("   [!] Erro: Campo CNPJ não apareceu.")
        return

    # Monitora Captcha
    if page.ele('text:Impedido por proteção Captcha', timeout=2):
        print("   [!] CAPTCHA DETECTADO! Resolva manualmente.")
        input("   -> Pressione ENTER aqui após resolver...")

    # Aguarda Login
    print("   -> Aguardando entrada no sistema...")
    contador = 0
    while contador < 60:
        if page.ele('text:Sair') or page.ele('text:Consulta') or "Identificacao" not in page.url:
            break
        time.sleep(1)
        contador += 1
        
    print("   -> Login detectado!")
    espera_aleatoria(2, 4)

    # 2. Navegação para Emissão
    print("   -> Buscando menu de Emissão...")
    try:
        if page.ele('text:Fechar', timeout=1): page.ele('text:Fechar').click()
    except: pass

    # Tenta clicar no menu
    if page.ele('text:Emitir Guia de Pagamento'):
        print("   -> Clicando no menu...")
        page.ele('text:Emitir Guia de Pagamento').click()
    else:
        print("   -> Menu não achado, indo por URL...")
        page.get("http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Emissao")

    espera_aleatoria(2, 3)

    # 3. Selecionar o Ano (CORRIGIDO)
    print(f"   -> Buscando ano {ANO_VIGENTE}...")
    
    try:
        # Procura o dropdown (select)
        dropdown = page.ele('tag:select') 
        
        if dropdown:
            print("   -> Menu de ano encontrado. Selecionando...")
            
            # --- CORREÇÃO AQUI ---
            # Removemos o "text=", passamos apenas o valor direto
            dropdown.select(ANO_VIGENTE)
            
            espera_aleatoria(1, 2)
            
            # Clica no botão OK (verde)
            botao_ok = page.ele('text:Ok') or page.ele('.btn-success')
            if botao_ok:
                botao_ok.click()
                print("   -> Botão OK clicado.")
            else:
                print("   [AVISO] Botão OK não encontrado.")
        else:
            print("   [AVISO] Dropdown de ano não encontrado.")

    except Exception as e:
        print(f"   [!] Erro na seleção do ano: {e}")

    # 4. Interação Humana
    print("\n   >>> SUA VEZ! <<<")
    print("   1. Marque o Mês.")
    print("   2. Gere e Baixe o DAS.")
    
    input("   >>> Pressione ENTER aqui APÓS o download terminar...")
    renomear_ultimo_arquivo(cnpj)

# === EXECUÇÃO ===
if __name__ == "__main__":
    if not os.path.exists(ARQUIVO_CLIENTES):
        print("ERRO: Planilha não encontrada.")
        exit()

    df = pd.read_excel(ARQUIVO_CLIENTES, dtype={'CNPJ': str})
    print(f"Carregado {len(df)} clientes.")
    
    print("Iniciando DrissionPage...")
    try:
        page = configurar_pagina()
    except Exception as e:
        print(f"Erro ao abrir navegador: {e}")
        exit()

    for index, row in df.iterrows():
        raw_cnpj = str(row['CNPJ'])
        cnpj_limpo = raw_cnpj.replace(".", "").replace("/", "").replace("-", "").strip()
        
        if len(cnpj_limpo) > 10:
            processar_cliente(page, cnpj_limpo)

    print("\nFim do processo.")