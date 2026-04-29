# UK Property Investment Calculator

A practical UK property investment calculator for quickly checking whether a deal makes sense on paper before wasting time, money, or energy on it.

This tool is designed for landlords, property investors, deal sourcers, and anyone analysing UK property opportunities such as:

- Buy-to-Let
- BRRR
- HMO-style investment deals
- Cashflow-focused property analysis
- Yield and return comparisons
- Basic investor-ready deal summaries

Built by **Marcin Sakowski**  
Website: [marcinsakowski.com](https://www.marcinsakowski.com)  
Live page: [Property Calculator](https://www.marcinsakowski.com/property-calculator)

---

## What This Tool Does

The calculator helps you analyse property deals by turning the messy numbers into a clearer picture.

Instead of guessing whether a deal is good, weak, or worth reviewing, you can enter the main figures and quickly see:

- Monthly cashflow
- Annual cashflow
- Gross yield
- Cash-on-cash return
- ROI / ROCE-style return metrics
- Estimated stamp duty / property tax logic
- Mortgage and finance cost estimates
- Refurb and holding cost impact
- Cash left in the deal
- Stress-test style outputs
- Share-ready summary text

It is not designed to replace professional financial, tax, legal, or mortgage advice. It is a practical first-pass calculator to help you spot obvious problems before going deeper.

---

## Why I Built It

Most property investors lose money before they even buy the property.

Not because the spreadsheet was complicated.

Because they ignored the basics:

- The rent was too optimistic
- The refurb budget was fantasy
- The mortgage cost was underestimated
- The cash left in the deal was too high
- The “deal” only worked when everything went perfectly

This calculator exists to make the numbers harder to lie about.

If a deal is weak, the calculator should make that obvious quickly.

---

## Versions Included

This repository includes multiple versions of the calculator:

| File | Purpose |
|---|---|
| `PropertyInvestmentCalculator.exe` | Windows executable version |
| `PropertyInvestmentCalculator.ps1` | PowerShell source version |
| `web-calculator.html` | Web / browser-based version |
| `LICENSE` | GPL-3.0 licence |

---

## Download

You can download the Windows executable directly from this repository:

[Download PropertyInvestmentCalculator.exe](./PropertyInvestmentCalculator.exe)

If Windows shows a security warning, that is normal for unsigned independent tools. Only run software you trust and always scan downloaded files if unsure.

---

## Run the PowerShell Version

If you prefer to run the source script directly:

```powershell
powershell -ExecutionPolicy Bypass -File .\PropertyInvestmentCalculator.ps1
