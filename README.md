# excel-para-investimentos

**Simulador de investimentos em FIIs no Excel — projeção de patrimônio, dividendos mensais e alocação por perfil de risco, com fórmulas nativas.**

![Excel](https://img.shields.io/badge/Excel-2016%2B-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![XLSX](https://img.shields.io/badge/arquivo-.xlsx-217346?style=flat-square)
![Openpyxl](https://img.shields.io/badge/gera%C3%A7%C3%A3o-openpyxl-4B8BBE?style=flat-square)
![Desafio DIO](https://img.shields.io/badge/Desafio-DIO-7B2CBF?style=flat-square)
![Categoria](https://img.shields.io/badge/categoria-Finan%C3%A7as%20Pessoais-2EA44F?style=flat-square)

Simulador de investimentos em Fundos de Investimento Imobiliário (FIIs) construído no Excel para o desafio DIO **"Ferramenta de Controle de Investimentos com Excel"**. Ele responde às perguntas que guiam qualquer plano de longo prazo: *quanto investir por mês? por quanto tempo? a que taxa de rendimento? qual patrimônio acumulado e quais dividendos mensais isso gera?* A engenharia está no que é entrega e o que é fórmula: **campos azuis são entradas do usuário** (Salário, Rendimento médio da carteira, Aporte mensal, Anos, Taxa mensal) e **campos verdes são cálculos nativos** — `FV()` para valor futuro, multiplicadores derivados para dividendos e total investido. Cenários de 2/5/10/20/30 anos e um Perfil de Alocação com menu suspenso de risco (Conservador/Moderado/Agressivo) que recalcula percentual e valor por tipo de FII. O arquivo é gerado programaticamente por `build_excel.py` com openpyxl, com `fix_alocacao.py` ajustando a aba de alocação — reprodutível, não planilha baixada pronta.

## 🇧🇷 Português

### Visão geral

O `Simulador_FII.xlsx` tem **duas abas**: o Simulador de projeção e o Perfil de Alocação. O fluxo de uso é direto: preencher os campos azuis, ler os verdes, escolher o perfil de risco.

**Convenção de cores (o contrato da planilha):**

- **Azul** — entrada do usuário: Salário, Rendimento médio da carteira (mensal, FIIs), Aporte mensal, Anos, Taxa mensal.
- **Verde** — fórmula (não editar): resultados calculados.

### Aba Simulador

Perguntas que a aba responde e as fórmulas que respondem:

| Pergunta | Fórmula (campo verde) |
|---|---|
| Quanto devo investir por mês? | `Sugestão de investimento = Salário * 30%` |
| Qual patrimônio acumulado no fim? | `Patrimônio = FV(taxa; anos*12; -aporte)` |
| Quanto vou receber de dividendos mensais? | `Dividendos mensais = patrimônio * rendimento` |
| Quanto sai do bolso no total? | `Total investido = aporte * anos * 12` |
| Quanto rendeu além dos aportes? | `Rendimento acumulado = patrimônio - TotalInvestido` |

A função `FV()` (valor futuro) é o motor da projeção: compostos os aportes mensais à taxa informada ao longo de `anos * 12` meses. O sinal negativo do aporte é a convenção da função (pagamento de saída) e devolve o valor futuro como número positivo.

### Cenários de horizonte

Projeções prontas para **2, 5, 10, 20 e 30 anos** — o mesmo aporte e a mesma taxa, cinco horizontes lado a lado, para ver o efeito dos juros compostos escala do tempo (a diferença entre 10 e 30 anos é tipicamente maior que a soma de todos os aportes).

### Perfil de Alocação

Bloco em colunas **K:L:M** com um **menu suspenso** (Data Validation) para o perfil de risco — **Conservador / Moderado / Agressivo**, selecionado na **célula L41**. A escolha recalcula automaticamente o **percentual e o valor mensal alocado por tipo de FII** (ex.: lajes corporativas, shoppings, galpões logísticos, fundos de papel): um perfil conservador concentra em tipos de menor volatilidade, um agressivo pesa nos de maior potencial de valorização.

### Arquivos do repositório

```
excel-para-investimentos/
├── Simulador_FII.xlsx          # Planilha: 2 abas (Simulador + Alocação)
├── build_excel.py              # Gera o arquivo via openpyxl
├── fix_alocacao.py             # Ajusta a aba de Perfil de Alocação
├── README.md
└── contexto do desafio.txt    # Enunciado original do desafio DIO
```

### Como usar

1. Abra `Simulador_FII.xlsx` no Excel (2016 ou superior recomendado).
2. Preencha **apenas os campos azuis**: salário, rendimento médio mensal esperado da carteira de FIIs, aporte mensal, horizonte em anos e taxa mensal.
3. Leia os resultados nos **campos verdes**: sugestão de investimento, patrimônio projetado, dividendos mensais, total investido e rendimento acumulado — para o horizonte escolhido e nos cenários de 2/5/10/20/30 anos.
4. Na aba de alocação, escolha seu **perfil de risco** (L41) e veja a distribuição por tipo de FII recalcular.

Para regenerar o arquivo do zero: `python build_excel.py` (requer `openpyxl`), depois `python fix_alocacao.py` se quiser reaplicar os ajustes da aba de alocação.

### Decisões de design

- **Fórmulas nativas, não valores colados**: `FV()` e os multiplicadores vivem nas células — mudar uma entrada recalcula tudo, e qualquer auditor pode clicar na célula e ver a lógica.
- **Contrato visual azul/verde**: separa sem ambiguidade o que é editável do que é calculado; reduz drasticamente o risco de usuário sobrescrever fórmula.
- **Menu suspenso com Data Validation**: perfil de risco como lista fechada elimina digitação livre e garante que a alocação por tipo de FII recalcule só com valores válidos.
- **Geração programática via openpyxl**: a planilha é artefato de código (`build_excel.py`), versionável e reproduzível — não um binário oposto de origem desconhecida.
- **Taxa mensal como entrada explícita**: FIIs pagam rendimento mensal (aluguéis); pedir taxa mensal em vez de anual alinha a entrada com a cadência real do ativo.

---

## 🇺🇸 English

### Overview

`Simulador_FII.xlsx` has **two sheets**: the projection Simulador and the Allocation Profile. Workflow: fill the blue cells, read the green ones, pick a risk profile.

**Color convention (the sheet's contract):**

- **Blue** — user input: Salary, average portfolio yield (monthly, FIIs), monthly contribution, Years, monthly rate.
- **Green** — formula (do not edit): computed results.

### Simulador sheet

Questions answered and the formulas answering them:

| Question | Formula (green field) |
|---|---|
| How much should I invest monthly? | `Suggested investment = Salary * 30%` |
| Final accumulated equity? | `Equity = FV(rate; years*12; -contribution)` |
| Monthly dividends at the end? | `Monthly dividends = equity * yield` |
| Total out of pocket? | `Total invested = contribution * years * 12` |
| Gains beyond contributions? | `Accumulated return = equity - TotalInvested` |

`FV()` (future value) drives the projection: monthly contributions compounded at the given rate over `years * 12` months. The negative contribution sign is the function's convention (outgoing payment), returning a positive future value.

### Horizon scenarios

Ready-made projections for **2, 5, 10, 20, and 30 years** — same contribution and rate, five horizons side by side, showing compounding's time scaling (the gap between 10 and 30 years is typically larger than the sum of all contributions).

### Allocation profile

Block in columns **K:L:M** with a **dropdown** (Data Validation) for the risk profile — **Conservative / Moderate / Aggressive**, selected at **cell L41**. The choice automatically recomputes the **percentage and monthly amount per FII type** (e.g., office buildings, malls, logistics warehouses, paper funds): a conservative profile concentrates in lower-volatility types, an aggressive one tilts toward higher appreciation potential.

### Repository files

```
excel-para-investimentos/
├── Simulador_FII.xlsx          # Workbook: 2 sheets (Simulador + Allocation)
├── build_excel.py              # Generates the file via openpyxl
├── fix_alocacao.py             # Patches the Allocation Profile sheet
├── README.md
└── contexto do desafio.txt    # Original DIO challenge statement (PT)
```

### How to use

1. Open `Simulador_FII.xlsx` in Excel (2016+ recommended).
2. Fill **only the blue cells**: salary, expected average monthly FII portfolio yield, monthly contribution, horizon in years, monthly rate.
3. Read results in the **green cells**: suggested investment, projected equity, monthly dividends, total invested, and accumulated return — for the chosen horizon and the 2/5/10/20/30-year scenarios.
4. On the allocation sheet, pick your **risk profile** (L41) and watch the per-FII-type distribution recompute.

To regenerate from scratch: `python build_excel.py` (requires `openpyxl`), then `python fix_alocacao.py` to re-apply allocation-sheet patches.

### Design decisions

- **Native formulas, not pasted values**: `FV()` and the multipliers live in the cells — changing one input recalculates everything, and any auditor can click a cell to inspect the logic.
- **Blue/green visual contract**: unambiguously separates editable from computed; drastically cuts the risk of users overwriting formulas.
- **Data Validation dropdown**: a closed list for the risk profile eliminates free typing and guarantees the per-FII allocation only recomputes on valid values.
- **Programmatic generation via openpyxl**: the workbook is a code artifact (`build_excel.py`) — versionable and reproducible, not an opaque binary of unknown origin.
- **Monthly rate as explicit input**: FIIs pay monthly income (rents); asking for the monthly rather than annual rate aligns the input with the asset's real cadence.

---

## Autor

**Kelvin Oliveira**

- GitHub: [KelvinOliveiraCode](https://github.com/KelvinOliveiraCode)
- LinkedIn: [kelvin-oliveira-code](https://www.linkedin.com/in/kelvin-oliveira-code/)

## Licença

Uso livre para estudo e adaptação. Trabalho desenvolvido para o desafio DIO "Ferramenta de Controle de Investimentos com Excel".
