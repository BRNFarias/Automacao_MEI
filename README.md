# 🤖 Automação de Emissão de Guias MEI (DAS) e Pagamento

Ferramenta de **RPA (Robotic Process Automation)** desenvolvida para escritórios de contabilidade e BPO Financeiro. O objetivo é automatizar o download das guias de pagamento (DAS) do Simples Nacional em lote e preparar os dados para pagamento no banco.

---

## 📋 Visão Geral

O sistema resolve dois grandes problemas manuais:

1. **Emissão**: Acessa o portal do Simples Nacional, baixa as guias (DAS) e organiza por CNPJ.
2. **Financeiro**: Lê os códigos de barras dos PDFs baixados e gera uma planilha formatada para **Pagamento em Lote** (Itaú, Inter, etc).

---

## 🚀 Funcionalidades Implementadas

- **Emissão em Lote**: Processa múltiplos CNPJs listados no ficheiro `clientes.xlsx`.
- **Cálculo Automático de Competência**: Identifica o mês de referência (mês anterior) sem necessidade de ajuste manual.
- **Validação Inteligente**:
  - Identifica e pula guias que já constam como **"Já existe pagamento"**.
  - Deteta e ignora empresas com situação **"BAIXADA"** ou **"ENCERRAMENTO"**.
- **Extração Bancária (Extrator Itaú)**: Lê os PDFs gerados, extrai a linha digitável (código de barras) e o valor, gerando um Excel pronto para importação no banco.

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.10+**
- **DrissionPage**: Para navegação web indetetável e robusta contra bloqueios.
- **Pandas & OpenPyXL**: Para manipulação de planilhas Excel.
- **pdfplumber**: Para leitura precisa dos dados dentro dos PDFs das guias.

---

## 📁 Estrutura do Projeto

```text
Automacao_MEI/
├── 1_baixar_guias.py          # Robô que acessa o site e baixa os PDFs
├── extrator_itau.py           # Robô que lê os PDFs e gera o Excel para pagamento
├── clientes.xlsx              # Ficheiro de entrada com a lista de CNPJs
├── das_baixados/              # Pasta onde os PDFs são salvos (gerada automaticamente)
├── lista_pagamentos_itau.xlsx # Relatório final gerado para o banco
├── requirements.txt           # Lista de dependências
└── README.md                  # Documentação
```

---

## ⚙️ Configuração e Instalação

1. **Instale as dependências**:

```bash
pip install -r requirements.txt
```

2. **Prepare a planilha de clientes**:

Crie ou edite o arquivo `clientes.xlsx` na raiz do projeto com, pelo menos, uma coluna chamada **CNPJ**.

---

## ▶️ Como Executar

### Passo 1: Baixar as Guias

Execute o script de download. O navegador abrirá, fará o login e baixará as guias automaticamente.

```bash
python 1_baixar_guias.py
```

> **Obs:** O robô pode solicitar interação manual em casos de CAPTCHA.

### Passo 2: Gerar Arquivo de Pagamento

Após o download das guias, execute o extrator para criar a lista de pagamento:

```bash
python extrator_itau.py
```

Isso irá gerar o ficheiro `lista_pagamentos_itau.xlsx` com os códigos de barras e valores extraídos dos PDFs.

---

## ⚠️ Aviso Legal

Ferramenta criada para fins de produtividade e gestão interna. O uso em portais governamentais deve seguir rigorosamente os termos de uso vigentes.

---

## 📄 Licença

Este projeto está licenciado sob a **MIT License**.

