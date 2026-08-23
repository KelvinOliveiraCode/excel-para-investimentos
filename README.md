# Simulador de Investimentos em Fundos Imobiliários (FII) no Excel

Repositório criado como entrega do desafio **"Ferramenta de Controle de Investimentos com Excel"** da [DIO](https://www.dio.me/). O objetivo é aplicar conceitos de Excel (fórmulas financeiras, referências, nomes definidos e formatação) na construção de uma planilha prática de **simulação de investimentos em Fundos Imobiliários (FIIs)**.

---

## 📌 Sobre o desafio

O laboratório propõe construir uma ferramenta que ajude o investidor a responder perguntas típicas:

- **Quanto devo investir por mês?**
- **Por quanto tempo (anos)?**
- **Qual a taxa de rendimento mensal?**
- **Quanto terei de patrimônio acumulado?**
- **Quanto receberei de dividendos mensais?**

A planilha automatiza cálculos complexos e gera uma visão clara do potencial de retorno, auxiliando decisões mais informadas.

---

## 📁 Arquivos do projeto

| Arquivo | Descrição |
|--------|-----------|
| `Simulador_FII.xlsx` | Planilha principal com o simulador (2 abas). |
| `fix_modelo.py` | Script que garante a coluna CHAVE da aba Perfis e ajusta a alocação por tipo de FII. |
| `build_excel.py` | Script Python (openpyxl) auxiliar de geração da planilha. |
| `contexto do desafio.txt` | Texto original do desafio. |
| `README.md` | Documentação do projeto. |

---

## 🛠️ Como usar a planilha

1. Abra `Simulador_FII.xlsx` no Microsoft Excel (ou LibreOffice Calc).
2. Na aba **Simulador**, edite os campos em **azul** (entradas do usuário):

| Campo | Significado |
|-------|-------------|
| Salário | Renda mensal do investidor |
| Rendimento médio da carteira | Dividendos médios mensais (% sobre o patrimônio) |
| Quanto investir por mês | Aporte mensal (R$) |
| Por quantos anos | Horizonte de investimento |
| Taxa de rendimento mensal | Valorização mensal estimada do patrimônio |

3. Os campos em **verde** são calculados automaticamente:

| Resultado | Fórmula |
|-----------|---------|
| Sugestão de investimento (30%) | `=Salário*30%` |
| Patrimônio acumulado | `=FV(taxa_mensal; qtd_anos*12; aporte*-1)` |
| Dividendos mensais | `=patrimonio*rendimento_carteira` |
| Total investido | `=aporte*qtd_anos*12` |
| Rendimento (juros) acumulado | `=patrimonio-TotalInvestido` |

4. A seção **Cenários** projeta o patrimônio e os dividendos em 2, 5, 10, 20 e 30 anos.
5. Na seção **Perfil de Alocação** (lateral direita, colunas K:L:M), use o **menu suspenso em "Perfil de risco" (célula L41)** para escolher entre Conservador, Moderado ou Agressivo. Ao alterar esse valor, a tabela recalcula automaticamente o **Percentual Sugerido** e o **Valor Mensal (R$)** de cada tipo de FII via `PROCV` contra a aba **Perfis**.
6. O gráfico **Evolução do Patrimônio, Investimento e Rendimento** (à direita) mostra, em 3 séries, a curva dos primeiros 60 meses: patrimônio acumulado (azul), total efetivamente investido (cinza) e rendimento/juros (verde).

---

## 🧮 Conceitos de Excel aplicados

- **Função financeira `FV`** (Valor Futuro): `FV(taxa; nper; pgto)` calcula o patrimônio acumulado de uma série de aportes.
- **`PROCV` (VLOOKUP)**: busca o percentual de alocação por tipo de FII a partir de uma chave `Perfil-Tipo`.
- **Nomes definidos**: `salario`, `rendimento_carteira`, `aporte`, `qtd_anos`, `taxa_mensal` e `patrimonio` tornam as fórmulas legíveis.
- **Formatação condicional de cores** e número (`R$` e `%`) para clareza visual.
- **Gráfico de linha** dinâmico alimentado por uma tabela de projeção mensal.

---

## 📊 Estrutura das abas

### Aba `Simulador`
- Cabeçalho e instruções.
- Bloco **Configurações do Investidor**.
- Bloco **Parâmetros da Simulação** (entradas + resultados).
- Bloco **Cenários** (2, 5, 10, 20 e 30 anos).
- Bloco **Projeção Mensal** (base do gráfico, 60 meses, com colunas Patrimônio, Total Investido e Rendimento).
- Bloco **Perfil de Alocação** à direita (PAPEL, TIJOLO, HÍBRIDOS, FOFs, DESENVOLVIMENTO, HOTELARIAS), com **menu suspenso de Perfil de Risco** que recalcula a alocação ao ser alterado.

### Aba `Perfis`
Tabela de referência com a alocação sugerida (%) por perfil de risco e tipo de FII, incluindo a coluna **CHAVE** (`Perfil-Tipo`) usada pelo `PROCV`.

---

## 🔁 Reprodutibilidade

A planilha foi ajustada programaticamente com `openpyxl` (garantia da coluna CHAVE de alocação):

```bash
pip install openpyxl
python fix_modelo.py
```

---

## ✅ Objetivos de aprendizagem atendidos

- [x] Criar ferramenta de simulação de investimentos em Excel.
- [x] Aplicar cálculos financeiros (rendimento mensal e dividendos).
- [x] Documentar processo técnico de forma clara e estruturada.
- [x] Utilizar o GitHub para compartilhar documentação técnica.

---

## ⚠️ Disclaimer

Esta ferramenta tem **finalidade educacional**. Os retornos simulados são estimativas baseadas nos parâmetros informados e não constituem recomendação de investimento.
