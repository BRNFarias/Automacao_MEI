# Automação de Emissão de Guias MEI (DAS) 🤖💸

Ferramenta de **RPA (Robotic Process Automation)** para escritórios de contabilidade e BPO Financeiro. O objetivo é automatizar o download das guias de pagamento (DAS) do Simples Nacional para diversos clientes MEI em lote.

## 🚀 Diferenciais Técnicos

O grande desafio foi contornar as defesas anti-robô do portal do Governo.

- **Solução:** Uso da biblioteca **DrissionPage** no lugar do Selenium.
- **Motivo:** DrissionPage controla o navegador via **CDP (Chrome DevTools Protocol)**, tornando a automação indetectável.
- **Benefício:** Digitação nativa e fluida, sem injeção de scripts.

## 📋 Funcionalidades

- **Leitura em lote:** Importa vários CNPJs a partir de um arquivo `clientes.xlsx`.
- **Navegação automática:** Acessa o PGMEI, preenche dados e segue para a emissão da guia.
- **Modo semi-automático:** Pausa em etapas críticas (Captcha/download) para intervenção humana.
- **Organização automática:** Renomeia o PDF baixado para o respectivo CNPJ.

## 🛠️ Tecnologias Utilizadas

- Python 3.10+
- DrissionPage
- Pandas
- OpenPyXL

## ⚙️ Configuração do Ambiente

### 1. Clone o repositório
```bash
git clone https://github.com/seu-usuario/automacao-mei.git
cd automacao-mei
```

### 2. Crie um ambiente virtual (opcional)
```bash
python -m venv venv

# Windows
./venv/Scripts/activate

# Linux/Mac
source venv/bin/activate
```

### 3. Instale as dependências
```bash
pip install -r requirements.txt
```

### 4. Prepare a planilha de clientes
Crie um arquivo **clientes.xlsx** na raiz com colunas:

| CNPJ           | NOME (Opcional) |
|----------------|------------------|
| 12345678000199 | Cliente A        |
| 98765432000100 | Cliente B        |

> Obs: Esse arquivo está no `.gitignore`.

## ▶️ Como Executar

1. Certifique-se de que o **Google Chrome esteja fechado**.
2. Execute o script:
```bash
python 1_baixar_guias.py
```
3. O navegador abrirá automaticamente.
4. O robô preencherá o CNPJ.
5. **Você interage apenas quando solicitado** (Captcha, confirmações, escolha de mês/ano).
6. Pressione **ENTER** no terminal para continuar para o próximo cliente.

Os arquivos serão salvos em **das_baixados/** já renomeados.

## 📁 Estrutura do Projeto
```
automacao-mei/
├── das_baixados/       # PDFs baixados
├── 1_baixar_guias.py   # Script principal
├── clientes.xlsx       # Base de dados (ignorado no Git)
├── requirements.txt    # Dependências
├── .gitignore
└── README.md
```

## ⚠️ Aviso Legal
Ferramenta criada para fins educacionais e de produtividade. O uso em portais governamentais deve seguir as normas e termos vigentes. O autor não é responsável por qualquer uso indevido.

**Desenvolvido por Seu Nome**

