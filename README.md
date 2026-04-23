# Excel Bonito

Aplicacao web para receber um arquivo Excel contábil, reorganizar os lancamentos em uma tabela padronizada e devolver uma versao limpa e mais profissional para download.

## O que o projeto faz

- recebe arquivos `.xls`, `.xlsx` e `.xlsm`
- remove linhas vazias
- detecta formatos genericos e layouts de Razao/Diario com varias colunas
- reconhece automaticamente os 3 modelos principais: Balancete, Livro Diario e Livro Razao
- identifica a coluna de data e o historico por cabecalho quando existir
- concatena linhas de continuacao na descricao do lancamento anterior
- separa valores em `Debito`, `Credito` e `Saldo`
- exporta uma nova planilha com colunas fixas `Data | Descricao | Debito | Credito | Saldo`
- aplica estilo visual no cabecalho, filtros, larguras e listras alternadas

## Como rodar

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Depois, abra no navegador:

`http://127.0.0.1:5000`

## Estrutura

- `app.py`: upload e download do arquivo
- `beautifier.py`: leitura, normalizacao e estilo do Excel
- `templates/index.html`: interface
- `static/styles.css`: visual da pagina
