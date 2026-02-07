# Apartment Investment Model

A simple Excel model to justify an apartment purchase: inputs, loan schedule, cash flows, and summary metrics (IRR, ROI, MOIC, breakeven).

## Quick start

1. Install Python dependency: `pip install openpyxl`
2. Run: `python build_model.py`
3. Open `apartment_investment_model.xlsx` and edit **Inputs** (purchase price, down payment, rent, etc.). All other sheets recalc automatically.

## Sheets

- **Inputs** — All assumptions (purchase price, down %, closing costs %, loan rate/term, monthly rent, rent growth, vacancy, op ex %, holding period, exit price growth). Edit here; everything else links to these cells.
- **Loan** — Monthly amortization (first 20 months shown). Annual debt service = 12 × monthly payment; used in Cash Flow.
- **Cash Flow** — Year 0 = equity out (down + closing). Years 1 to holding: gross rent (with growth), vacancy, effective rent, op ex, NOI, debt service, levered CF, sale proceeds (only in exit year), total CF, cumulative CF.
- **Summary** — Exit sale price, loan balance at exit, net sale proceeds; equity invested; total distributions; **MOIC**, **ROI %**, **IRR %**, **breakeven year**.

## Metrics

- **IRR** — Internal rate of return on equity (year 0 outflow + annual total CF through exit).
- **ROI** — (Total distributions − Equity) / Equity.
- **MOIC** — Total distributions / Equity invested.
- **Breakeven** — First year when cumulative cash flow turns non‑negative.

## Currency

Use one currency consistently (e.g. MXN or USD). Defaults are in local currency; change labels in Inputs if you prefer.
