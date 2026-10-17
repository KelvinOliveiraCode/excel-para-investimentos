# Real Estate Investment Fund (FII) Investment Simulator in Excel

Repository created as the deliverable for the challenge **"Investment Control Tool with Excel"** from [DIO](https://www.dio.me/). The objective is to apply Excel concepts (financial formulas, references, defined names, and formatting) in building a practical spreadsheet for **simulating investments in Real Estate Investment Funds (FIIs)**.

---

## About the challenge

The lab proposes building a tool that helps the investor answer typical questions:

- How much should I invest per month?
- For how long (years)?
- What is the monthly return rate?
- How much accumulated equity will I have?
- How much will I receive in monthly dividends?

The spreadsheet automates complex calculations and generates a clear view of potential return, assisting more informed decisions.

---

## Project files

| File | Description |
|------|-------------|
| `Simulador_FII.xlsx` | Main spreadsheet with the simulator (2 sheets). |
| `fix_alocacao.py` | Script that ensures the CHAVE column of the Perfis sheet and adjusts allocation by FII type. |
| `build_excel.py` | Auxiliary Python (openpyxl) script for generating the spreadsheet. |
| `contexto do desafio.txt` | Original challenge text. |
| `README.md` | Project documentation. |

---

## How to use the spreadsheet

1. Open `Simulador_FII.xlsx` in Microsoft Excel (or LibreOffice Calc).
2. On the **Simulador** sheet, edit the fields in **blue** (user inputs):

| Field | Meaning |
|-------|---------|
| Salary | Investor's monthly income |
| Average portfolio return | Average monthly dividends (% of equity) |
| How much to invest per month | Monthly contribution (R$) |
| For how many years | Investment horizon |
| Monthly return rate | Estimated monthly equity appreciation |

3. The fields in **green** are calculated automatically:

| Result | Formula |
|--------|---------|
| Investment suggestion (30%) | `=Salario*30%` |
| Accumulated equity | `=FV(taxa_mensal; qtd_anos*12; aporte*-1)` |
| Monthly dividends | `=patrimonio*rendimento_carteira` |
| Total invested | `=aporte*qtd_anos*12` |
| Accumulated return (interest) | `=patrimonio-TotalInvestido` |

4. The **Scenarios** section projects equity and dividends over 2, 5, 10, 20, and 30 years.
5. In the **Allocation Profile** section (right side, columns K:L:M), use the **dropdown in "Risk Profile" (cell L41)** to choose Conservative, Moderate, or Aggressive. When changing this value, the table automatically recalculates the **Suggested Percentage** and **Monthly Value (R$)** of each FII type via `PROCV` against the **Perfis** sheet.
6. The chart **Equity, Investment, and Return Evolution** (right side) shows, in 3 series, the curve of the first 60 months: accumulated equity (blue), total effectively invested (gray), and return/interest (green).

---

## Excel concepts applied

- Financial function `FV` (Future Value): `FV(rate; nper; pmt)` calculates accumulated equity from a series of contributions.
- `PROCV` (VLOOKUP): looks up allocation percentage by FII type from a `Profile-Type` key.
- Defined names: `salario`, `rendimento_carteira`, `aporte`, `qtd_anos`, `taxa_mensal` and `patrimonio` make formulas readable.
- Conditional color formatting and numbers (`R$` and `%`) for visual clarity.
- Line chart fed by a monthly projection table.

---

## Sheet structure

### `Simulador` sheet

- Header and instructions.
- **Investor Settings** block.
- **Simulation Parameters** block (inputs + results).
- **Scenarios** block (2, 5, 10, 20, and 30 years).
- **Monthly Projection** block (chart basis, 60 months, with Equity, Total Invested, and Return columns).
- **Allocation Profile** block on the right (PAPEL, TIJOLO, HÍBRIDOS, FOFs, DESENVOLVIMENTO, HOTELARIAS), with **Risk Profile dropdown** that recalculates allocation when changed.

### `Perfis` sheet

Reference table with suggested allocation (%) by risk profile and FII type, including the **CHAVE** column (`Perfil-Tipo`) used by `PROCV`.

---

## Reproducibility

The spreadsheet was adjusted programmatically with `openpyxl` (allocation CHAVE column guarantee):

```bash
pip install openpyxl
python fix_alocacao.py
```

---

## Learning objectives met

- [x] Create an investment simulation tool in Excel.
- [x] Apply financial calculations (monthly return and dividends).
- [x] Document technical process clearly and structured.
- [x] Use GitHub to share technical documentation.

---

## Disclaimer

This tool has **educational purposes**. Simulated returns are estimates based on informed parameters and do not constitute investment advice.
