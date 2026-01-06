# Automação MEI

Projeto em Python para automatizar tarefas financeiras de um MEI, com foco na extração e organização de extratos bancários do Itaú.

## Visão geral

Este projeto foi criado para reduzir trabalho manual na gestão financeira do MEI, automatizando a coleta e o tratamento de dados bancários.

Atualmente, o foco principal é a extração de extratos do Itaú e a transformação dessas informações em dados estruturados para análise e controle financeiro.

## Funcionalidades implementadas

- Extração automatizada de extratos bancários do Itaú
- Coleta de lançamentos financeiros (data, descrição, valor)
- Organização dos dados em formato estruturado
- Preparação dos dados para uso contábil ou financeiro
- Automação de tarefas repetitivas do dia a dia do MEI

## Extrator Itaú

O arquivo `extrator_itau.py` é responsável por:

- Acessar o sistema do Itaú
- Navegar até a área de extratos
- Extrair os lançamentos financeiros do período disponível
- Processar os dados extraídos
- Gerar saída organizada (ex: CSV ou outro formato estruturado)

Esse script é o núcleo da automação financeira do projeto.

## Requisitos

- Python 3.10 ou superior
- Dependências listadas no arquivo `requirements.txt`

Instalação das dependências:

```bash
pip install -r requirements.txt
```

## Como executar

Execute o script principal com:

```bash
python extrator_itau.py
```

Caso o script utilize parâmetros adicionais (como datas, conta ou credenciais), ajuste conforme a implementação no arquivo.

## Estrutura do projeto

```text
Automacao_MEI/
├── extrator_itau.py
├── requirements.txt
├── README.md
└── outros arquivos auxiliares
```

## Objetivo do projeto

- Facilitar a organização financeira do MEI
- Evitar erros manuais na leitura de extratos
- Economizar tempo com automação
- Servir como base para futuras automações fiscais e contábeis

## Próximos passos

- Melhorar tratamento de erros
- Adicionar logs mais detalhados
- Suporte a outros bancos
- Integração com planilhas ou banco de dados

## Licença

Este projeto está sob a licença MIT.
