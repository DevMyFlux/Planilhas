# PROMPT DE VALIDAÇÃO — KIT CONTÁBIL TOTVS → HCN

Você receberá **dois arquivos Excel** por módulo:
- **BASE** — saída do sistema (arquivo `_organizado_*.xlsx`)
- **GAB** — gabarito do funcionário (`*_plt.xlsx` ou `*_hef.xlsx`)

Gere um **Relatório de Compatibilidade** com a estrutura abaixo.
Siga à risca as regras de comparação para cada módulo.

---

## REGRAS GERAIS

- Ignore diferenças de formatação (número de casas decimais, separador BR vs US,
  maiúsculas/minúsculas). Compare **valores numéricos**, não strings.
- Arredonde valores monetários a 2 casas decimais antes de comparar.
- Se um arquivo tiver coluna extra que o outro não tem (ex: `Mov` no Diário), ignore-a
  na comparação de conteúdo — não é um erro.
- **Nunca compare por posição de linha quando a chave de identidade puder ser usada.**
  Comparar a linha N de um arquivo com a linha N do outro só é válido quando a ordem
  é garantidamente idêntica (ver Razão abaixo).

---

## MÓDULO 0300 — BALANCETE

### Chave de identidade: `ContaContabil`

Antes de qualquer comparação de valores, **alinhe os dois arquivos por `ContaContabil`**.
Toda comparação subsequente — SaldoAnterior, Débito, Crédito, SaldoAtual, flags S/N,
campos NaN vs preenchido — deve ser feita **entre linhas com o mesmo ContaContabil**,
nunca entre a linha N do GAB e a linha N da BASE.

Uma conta ausente nunca deve gerar divergências reportadas nas linhas seguintes a ela.
Se isso ocorrer, é sinal de que a comparação está sendo feita por posição: pare e
refaça por chave.

### Dimensões

Contar linhas de dados (excluindo cabeçalho) em cada arquivo.
Reportar: `GAB = X linhas | BASE = Y linhas` e ✅ se iguais, ⚠️ se diferentes.

### Comparação por chave

Após alinhar por `ContaContabil`, reportar **separadamente**:

**(a) Contas presentes em ambos com valores divergentes**
Para cada `ContaContabil` comum, comparar célula a célula:
`SaldoAnterior`, `Débito`, `Crédito`, `SaldoAtual`, `Estoque`, `ContaFinanceira`,
`ReservaDeContingência`, `AtivoImob`, `DepreciacaoAcumulativa`.
Listar cada célula divergente: `Conta X, coluna Y: GAB=valor / BASE=valor`.

**(b) Contas presentes só no GAB — faltando no BASE**
Lista de `ContaContabil` presentes no GAB mas ausentes no BASE.
Classificar como 🔴 ALTO (conta com valores) ou 🟡 BAIXO (conta zerada).

**(c) Contas presentes só no BASE — extras não esperadas**
Lista de `ContaContabil` presentes no BASE mas ausentes no GAB.

### Totais financeiros

Somar `Débito` e `Crédito` de cada arquivo independentemente e comparar.

### Resumo do Balancete

```
Dimensões:   GAB=X  BASE=Y  [✅/⚠️]
Totais Deb:  GAB=X  BASE=Y  [✅/⚠️]
Totais Cre:  GAB=X  BASE=Y  [✅/⚠️]
Contas comuns com diffs: N
Contas só no GAB: N
Contas só no BASE: N
```

---

## MÓDULO 1600 — LIVRO DIÁRIO

### Chave de identidade: `(DATA, CLASSIFICAÇÃO, DÉBITO, CRÉDITO)`

O Diário **não tem ordem canônica obrigatória** — apenas o conjunto de lançamentos
precisa ser idêntico. Não assuma que o par na posição K do BASE corresponde ao par
na posição K do GAB. Se a ordem dos lançamentos diferir, todos os pares a partir do
primeiro desalinhamento seriam falsos positivos se comparados por posição.

### Dimensões

Contar linhas de dados (excluindo cabeçalho) em cada arquivo.
Reportar: `GAB = X linhas | BASE = Y linhas` e ✅ se iguais, ⚠️ se diferentes.

### Totais financeiros

Somar colunas `DÉBITO` e `CRÉDITO` de cada arquivo e comparar.

### Comparação por chave de conteúdo

Para cada linha de cada arquivo, construir a chave:
`(DATA, CLASSIFICAÇÃO, round(DÉBITO, 2), round(CRÉDITO, 2))`

Comparar os dois arquivos pelo **conjunto de chaves**:

**(a) Lançamentos comuns — verificar HISTÓRICO**
Chaves presentes em ambos: lançamento confirmado.
Dentro dessas chaves, comparar a coluna `HISTÓRICO`/`DESCRIÇÃO` entre as duas
ocorrências. Listar divergências de texto (ignorar diferenças de espaço duplo vs simples
se o contexto indicar que o sistema normaliza espaços).

**(b) Lançamentos só no GAB — faltando no BASE** 🔴 ALTO
Chaves presentes no GAB mas ausentes no BASE. Lista com DATA, CLASSIFICAÇÃO e valor.

**(c) Lançamentos só no BASE — extras não esperados** 🔴 ALTO
Chaves presentes no BASE mas ausentes no GAB.

### Verificação de ordem interna débito/crédito

Dentro de cada par débito/crédito vinculado (mesmo número de movimento `Mov` quando
disponível, ou contrapartida identificável), verificar se a ordem está correta:
- Conta de PASSIVO (2.x) ou ATIVO (1.x) deve vir **antes** da conta de DESPESA (3.x).
- Reportar cada par com ordem invertida: `Par Mov=NNNN: GAB=[1.x, 3.x] BASE=[3.x, 1.x]`.

Classificar cada par como:
- ✅ **Idêntico** — conteúdo e ordem corretos
- 🟡 **Invertido** — mesmo conteúdo, apenas ordem interna trocada
- 🔴 **Diferente** — chave ausente em um dos arquivos (nunca por posição diferente)

### Resumo do Diário

```
Dimensões:          GAB=X  BASE=Y  [✅/⚠️]
Total DÉBITO:       GAB=X  BASE=Y  [✅/⚠️]
Total CRÉDITO:      GAB=X  BASE=Y  [✅/⚠️]
Lançamentos comuns: N (X% do total)
Só no GAB (falta):  N
Só no BASE (extra): N
Histórico divergente: N
Pares com ordem invertida: N
```

---

## MÓDULO 1700 — LIVRO RAZÃO

### Comparação por posição de linha

O Razão é ordenado por `Conta Analítica + Data` e essa ordem é fixa e determinística
em ambos os arquivos. A comparação **por posição de linha é válida aqui**.

### Dimensões

Contar linhas de dados (excluindo cabeçalho e linhas de cabeçalho de conta).
Reportar: `GAB = X linhas | BASE = Y linhas` e ✅ se iguais, ⚠️ se diferentes.

### Comparação célula a célula por posição

Comparar linha N do GAB com linha N do BASE para todas as colunas comuns.
Se as dimensões diferirem, alinhar até a menor dimensão e reportar as linhas extras.

Para divergências no `HISTÓRICO`: verificar se o texto do BASE é **mais completo**
que o do GAB (o GAB pode ter texto truncado pelo funcionário — isso não é erro do
sistema). Classificar como:
- **GAB truncado** (BASE tem texto mais completo) → sistema correto, não é erro
- **BASE truncado** (GAB tem texto mais completo) → erro do sistema 🔴
- **Texto trocado** (completamente diferente) → erro do sistema 🔴

### Resumo do Razão

```
Dimensões:             GAB=X  BASE=Y  [✅/⚠️]
Histórico GAB truncado (sistema mais completo): N  [não são erros]
Histórico BASE truncado (erro): N
Texto trocado (erro): N
Outros diffs: N
```

---

## ENTREGA — RELATÓRIO FINAL

Produza um relatório consolidado com:

### Cabeçalho
```
RELATÓRIO DE COMPATIBILIDADE
KIT Contábil MM/AAAA — [Unidade]
Comparativo: Organizado (sistema) × Gabarito (funcionário)
```

### Tabela de Status Geral

| Módulo | Status | Resumo |
|--------|--------|--------|
| Razão (1700) | ✅/⚠️/🔴 | N diffs — [breve descrição] |
| Diário (1600) | ✅/⚠️/🔴 | N lançamentos, N faltando, N extras |
| Balancete (0300) | ✅/⚠️/🔴 | N contas divergentes, N faltando, N extras |

### Detalhamento por módulo

Para cada módulo, uma seção com:
1. Dimensões e totais financeiros
2. Lista de divergências reais (por chave, não por posição)
3. Classificação de severidade (🔴 ALTO / 🟠 MÉDIO / 🟡 BAIXO)
4. Conclusão: o sistema está correto / precisa de correção

### Abas da planilha de entrega (opcional)

Se entregar planilha:
- *1700 — Razão*: comparação posicional, divergências de histórico com classificação
- *1600 — Diário*: quatro tabelas:
  - **Tab 1** — Lançamentos com chave comum mas histórico diferente
  - **Tab 2** — Lançamentos só no GAB (faltando no sistema)
  - **Tab 3** — Lançamentos só no BASE (extras inesperados)
  - **Tab 4** — Pares com ordem interna débito/crédito invertida
- *0300 — Balancete*: três seções:
  - **Sec A** — Contas comuns com valores divergentes (por chave ContaContabil)
  - **Sec B** — Contas só no GAB (faltando no sistema)
  - **Sec C** — Contas só no BASE (extras inesperados)

---

## CHECKLIST ANTI-FALSO-POSITIVO

Antes de reportar uma divergência como erro, verificar:

- [ ] A comparação foi feita por chave (não por posição)?
- [ ] Se uma conta está "faltando", verificar se não é apenas desalinhamento posicional
      (a mesma ContaContabil existe em outra linha)?
- [ ] Se um lançamento está "faltando" no Diário, verificar se não é apenas
      ordem diferente (a mesma chave DATA+CLASSIFICAÇÃO+DÉBITO+CRÉDITO existe
      em outra posição)?
- [ ] Diferença de histórico: o BASE tem texto mais completo que o GAB
      (não é erro — é melhoria)?
- [ ] Coluna `Mov` extra no Diário: não é erro, é dado adicional do sistema.
