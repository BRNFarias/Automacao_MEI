import os
import time
import pandas as pd
from DrissionPage import ChromiumPage, ChromiumOptions

# === CONFIGURAÇÕES ===
PASTA_PROJETO = os.getcwd()
PASTA_DOWNLOAD = os.path.join(PASTA_PROJETO, "das_baixados")
ARQUIVO_CLIENTES = "clientes.xlsx"
ANO_VIGENTE = "2025"

# Cria pasta de download se não existir
if not os.path.exists(PASTA_DOWNLOAD):
    os.makedirs(PASTA_DOWNLOAD)

def configurar_pagina():
    # Configura o navegador
    co = ChromiumOptions()
    
    # --- CORREÇÃO AQUI (Tiramos o 's' do final) ---
    co.set_download_path(PASTA_DOWNLOAD)
    
    # Tenta conectar ao Chrome instalado
    page = ChromiumPage(co)
    return page

def renomear_ultimo_arquivo(cnpj):
    # Espera um pouco para garantir que o download acabou
    time.sleep(3) 
    
    arquivos = [os.path.join(PASTA_DOWNLOAD, f) for f in os.listdir(PASTA_DOWNLOAD) if f.endswith(".pdf")]
    
    if not arquivos:
        print("   [!] Nenhum PDF novo encontrado na pasta.")
        return
        
    # Pega o arquivo mais recente
    ultimo_arquivo = max(arquivos, key=os.path.getctime)
    novo_nome = os.path.join(PASTA_DOWNLOAD, f"{cnpj}.pdf")
    
    # Se já existe (ex: teste anterior), apaga
    if os.path.exists(novo_nome):
        try: os.remove(novo_nome)
        except: pass
        
    try:
        os.rename(ultimo_arquivo, novo_nome)
        print(f"   [OK] PDF salvo como: {cnpj}.pdf")
    except Exception as e:
        print(f"   [ERRO] Falha ao renomear: {e}")

def processar_cliente(page, cnpj):
    print(f"\n--- Processando CNPJ: {cnpj} ---")
    
    # 1. Acessa o site
    page.get("http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Identificacao")
    
    try:
        # 2. Preenche o CNPJ
        # O DrissionPage limpa e digita nativamente
        if page.ele('#cnpj'):
            page.ele('#cnpj').clear()
            page.ele('#cnpj').input(cnpj)
            print("   -> CNPJ digitado.")
            
            # 3. Clica em Continuar
            page.ele('#continuar').click()
        else:
            print("   [!] Campo CNPJ não encontrado. O site carregou?")
            return

        # Verifica se pediu Captcha
        # O timeout=2 significa que ele olha rapidinho se apareceu a mensagem
        if page.ele('text:Impedido por proteção Captcha', timeout=2):
            print("   [!] CAPTCHA DETECTADO!")
            print("   -> Resolva manualmente no navegador agora.")
            input("   -> Pressione ENTER aqui depois de resolver...")
        
        # Espera login (aguarda a URL mudar ou o menu aparecer)
        print("   -> Aguardando login...")
        page.wait.url_change('Identificacao', timeout=300)
        print("   -> Logado!")
        
        # 4. Navega para Emissão
        time.sleep(1)
        # Tenta clicar no menu se ele existir, senão vai pela URL direta
        if page.ele('tag:a@@text():Emitir Guia de Pagamento'):
            page.ele('tag:a@@text():Emitir Guia de Pagamento').click()
        else:
            page.get("http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Emissao")

        # 5. Seleciona o Ano
        print(f"   -> Tentando selecionar {ANO_VIGENTE}...")
        try:
            # Procura o texto do ano na tela e clica
            if page.ele(f'text:{ANO_VIGENTE}'):
                page.ele(f'text:{ANO_VIGENTE}').click()
            else:
                print(f"   [AVISO] Ano {ANO_VIGENTE} não encontrado. Selecione manualmente.")
        except:
            pass

        # 6. Intervenção Final
        print("\n   >>> AÇÃO MANUAL NECESSÁRIA <<<")
        print("   1. Marque o Mês.")
        print("   2. Clique em 'Apurar/Gerar DAS'.")
        print("   3. Baixe o PDF.")
        
        input("   >>> Pressione ENTER aqui DEPOIS que o download terminar...")
        
        renomear_ultimo_arquivo(cnpj)
        
    except Exception as e:
        print(f"   [ERRO CRÍTICO] {e}")

# === EXECUÇÃO ===
if __name__ == "__main__":
    # Carrega planilha
    if not os.path.exists(ARQUIVO_CLIENTES):
        print("ERRO: Planilha não encontrada.")
        exit()

    df = pd.read_excel(ARQUIVO_CLIENTES, dtype={'CNPJ': str})
    print(f"Carregado {len(df)} clientes.")
    
    # Importante: Feche o Chrome antes de rodar isso
    print("Iniciando DrissionPage...")
    try:
        page = configurar_pagina()
    except Exception as e:
        print(f"Erro ao abrir navegador: {e}")
        print("DICA: Feche todas as janelas do Chrome antes de rodar o script.")
        exit()

    for index, row in df.iterrows():
        # Limpeza do CNPJ
        raw_cnpj = str(row['CNPJ'])
        cnpj_limpo = raw_cnpj.replace(".", "").replace("/", "").replace("-", "").strip()
        
        if len(cnpj_limpo) > 10: # Validação básica
            processar_cliente(page, cnpj_limpo)

    print("\nFim do processo.")