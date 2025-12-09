import os
import time
import random
import pandas as pd
import shutil
import datetime
from DrissionPage import ChromiumPage, ChromiumOptions

# === 1. CONFIGURAÇÕES GERAIS ===

PASTA_PROJETO = os.getcwd()
PASTA_DOWNLOAD = os.path.join(PASTA_PROJETO, "das_baixados")
ARQUIVO_CLIENTES = "clientes.xlsx"

# Garante que a pasta existe
if not os.path.exists(PASTA_DOWNLOAD):
    os.makedirs(PASTA_DOWNLOAD)

# === 2. FUNÇÕES AUXILIARES ===

def configurar_pagina():
    co = ChromiumOptions()
    co.set_download_path(PASTA_DOWNLOAD)
    
    # Configurações para evitar popups de download
    co.set_pref('profile.default_content_settings.popups', 0)
    co.set_pref('download.prompt_for_download', False)
    co.set_pref('download.directory_upgrade', True)
    
    page = ChromiumPage(co)
    return page

def espera_aleatoria(min=1, max=3):
    time.sleep(random.uniform(min, max))

def obter_competencia_atual():
    hoje = datetime.date.today()
    if hoje.month == 1:
        mes_ref = 12
        ano_ref = hoje.year - 1
    else:
        mes_ref = hoje.month - 1
        ano_ref = hoje.year

    nomes_meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    return nomes_meses[mes_ref], str(ano_ref)

def buscar_e_renomear(cnpj, tempo_do_clique):
    """
    Procura um arquivo criado APÓS o tempo_do_clique.
    """
    print("   -> Monitorando chegada do arquivo...")
    
    pasta_projeto = PASTA_DOWNLOAD
    pasta_windows = os.path.join(os.path.expanduser("~"), "Downloads")
    
    arquivo_encontrado = None
    origem = None
    
    # Tenta encontrar o arquivo por 30 segundos
    tentativas = 0
    while tentativas < 30:
        time.sleep(1)
        
        # 1. Verifica na pasta do PROJETO
        for f in os.listdir(pasta_projeto):
            if f.endswith(".pdf"):
                caminho = os.path.join(pasta_projeto, f)
                # Verifica se o arquivo é MAIS NOVO que o clique
                if os.path.getctime(caminho) > tempo_do_clique:
                    arquivo_encontrado = caminho
                    origem = "projeto"
                    break
        
        # 2. Se não achou, verifica na pasta DOWNLOADS (Backup)
        if not arquivo_encontrado and os.path.exists(pasta_windows):
            for f in os.listdir(pasta_windows):
                if f.endswith(".pdf"):
                    caminho = os.path.join(pasta_windows, f)
                    if os.path.getctime(caminho) > tempo_do_clique:
                        arquivo_encontrado = caminho
                        origem = "windows"
                        break
        
        if arquivo_encontrado:
            break
            
        tentativas += 1
        if tentativas % 5 == 0:
            print(f"   ...aguardando ({tentativas}s)...")

    if not arquivo_encontrado:
        print("   [ERRO] Timeout: O arquivo não apareceu em 30 segundos.")
        return

    time.sleep(1) # Segurança final de escrita em disco

    novo_nome = os.path.join(pasta_projeto, f"{cnpj}.pdf")
    
    if os.path.exists(novo_nome):
        try: os.remove(novo_nome)
        except: pass
        
    try:
        if origem == "windows":
            shutil.move(arquivo_encontrado, novo_nome)
            print(f"   [SUCESSO] Recuperado de Downloads e salvo como: {cnpj}.pdf")
        else:
            os.rename(arquivo_encontrado, novo_nome)
            print(f"   [SUCESSO] Salvo na pasta correta como: {cnpj}.pdf")
            
    except Exception as e:
        print(f"   [ERRO] Falha ao mover/renomear arquivo: {e}")

# === 3. LÓGICA DO ROBÔ ===

def processar_cliente(page, cnpj):
    mes_alvo, ano_alvo = obter_competencia_atual()
    
    print(f"\n========================================")
    print(f"CNPJ: {cnpj} | Competência: {mes_alvo}/{ano_alvo}")
    
    url_base = "http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Identificacao"
    page.get(url_base)
    
    # Limpeza de Sessão
    if page.ele('text:Sair') or page.ele('text:Trocar usuário'):
        page.ele('text:Sair').click()
        espera_aleatoria(1, 1.5)
        page.get(url_base)

    # Login
    if page.ele('#cnpj'):
        page.ele('#cnpj').clear()
        page.ele('#cnpj').input(cnpj)
        espera_aleatoria(0.5, 1)
        if page.ele('#continuar'):
            page.ele('#continuar').click()
    else:
        print("   [ERRO] Campo CNPJ não encontrado.")
        return

    # Captcha
    if page.ele('text:Impedido por proteção Captcha', timeout=2):
        print("   [!] RESOLVA O CAPTCHA MANUALMENTE!")
        input("   -> Pressione ENTER aqui no terminal após resolver...")

    # Verifica Login OK
    contador = 0
    logado = False
    while contador < 30:
        if page.ele('text:Sair') or "Identificacao" not in page.url:
            logado = True
            break
        time.sleep(1)
        contador += 1
    
    if not logado:
        print("   [ERRO] Falha no login (Timeout).")
        return
        
    # Verifica Empresa Baixada
    texto_pagina = page.html.upper()
    avisos = ['SITUAÇÃO CADASTRAL: BAIXADA', 'CNPJ BAIXADO', 'EXTINÇÃO POR ENCERRAMENTO', 'SITUAÇÃO: NULA']
    for aviso in avisos:
        if aviso in texto_pagina:
            print(f"   [ALERTA] Empresa Baixada/Encerrada. Pulando.")
            return

    espera_aleatoria(1, 2)
    
    # Navegação
    if page.ele('text:Fechar', timeout=1): page.ele('text:Fechar').click()
    
    if page.ele('text:Emitir Guia de Pagamento'):
        page.ele('text:Emitir Guia de Pagamento').click()
    else:
        page.get("http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Emissao")

    espera_aleatoria(2, 3)

    # Seleciona Ano
    try:
        dropdown = page.ele('tag:select')
        if dropdown:
            dropdown.select(ano_alvo)
            espera_aleatoria(1, 1.5)
            btn = page.ele('text:Ok') or page.ele('.btn-success')
            if btn: btn.click()
    except: pass

    espera_aleatoria(1, 2)

    # Seleciona Mês
    try:
        elm_mes = page.ele(f'text:{mes_alvo}')
        if elm_mes:
            # Clica Checkbox
            elm_mes.parent('tag:tr').ele('tag:input').click()
            espera_aleatoria(0.5, 1)
            
            # Clica Apurar
            btn_apurar = page.ele('text:Apurar') or page.ele('text:Gerar DAS') or page.ele('#btnEmitirDas')
            if btn_apurar:
                btn_apurar.click()
                
                # Verifica "Já Pago"
                time.sleep(2)
                if page.ele('text:Já existe pagamento') or page.ele('text:Não será gerado DAS'):
                    print(f"   [AVISO] Já pago ({mes_alvo}). Pulando.")
                    return

                # Imprimir (Download)
                btn_imprimir = page.ele('text:Imprimir', timeout=10) or page.ele('text:Visualizar PDF', timeout=10)
                
                if btn_imprimir:
                    tempo_antes_do_clique = time.time() # Marca a hora
                    btn_imprimir.click() 
                    buscar_e_renomear(cnpj, tempo_antes_do_clique)
                else:
                    print("   [ERRO] Botão Imprimir não apareceu.")
            else:
                print("   [ERRO] Botão Apurar não encontrado.")
        else:
            print(f"   [ALERTA] Mês {mes_alvo} indisponível.")
    except Exception as e:
        print(f"   [ERRO] Falha na emissão: {e}")

# === 4. EXECUÇÃO PRINCIPAL ===

if __name__ == "__main__":
    print("=== ROBÔ DAS AUTOMÁTICO - VERSÃO FINAL ===")
    
    if not os.path.exists(ARQUIVO_CLIENTES):
        print(f"ERRO: Arquivo '{ARQUIVO_CLIENTES}' não encontrado.")
        exit()

    df = pd.read_excel(ARQUIVO_CLIENTES, dtype={'CNPJ': str})
    print(f"Clientes na fila: {len(df)}")
    
    # Inicializa a variável page fora do try para garantir acesso no finally
    page = None

    try:
        page = configurar_pagina()
        
        for index, row in df.iterrows():
            cnpj_limpo = str(row['CNPJ']).replace(".", "").replace("/", "").replace("-", "").strip()
            
            if len(cnpj_limpo) >= 14:
                processar_cliente(page, cnpj_limpo)
            else:
                print(f"[PULAR] CNPJ inválido linha {index+2}")
                
    except Exception as e:
        print(f"\nERRO GERAL DO SISTEMA: {e}")
        
    finally:
        # === AQUI ESTÁ A LÓGICA DE FECHAMENTO ===
        print("\n=== TODOS OS PROCESSOS FINALIZADOS ===")
        if page:
            print("Fechando navegador...")
            page.quit() # Fecha o navegador
            print("Navegador fechado.")