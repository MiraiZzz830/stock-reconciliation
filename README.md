# Stock Reconciliation Skill

> **Deterministic inventory and catalog reconciliation between POS exports and online channels (Marketplaces & Web Stores) with 0 mistakes.**

[![Skill](https://img.shields.io/badge/Antigravity-Skill-blue.svg)](https://github.com/MiraiZzz830/stock-reconciliation)
[![Status](https://img.shields.io/badge/Status-Production--Ready-green.svg)](https://github.com/MiraiZzz830/stock-reconciliation)
[![Cross-Platform](https://img.shields.io/badge/Engine-Python%20%7C%20C%23%20.NET-brightgreen.svg)](#running-the-reconciliation)

---

## Overview

The **Stock Reconciliation Skill** provides deterministic matching and discrepancy auditing between:
- **POS Systems**: SQL POS, Autocount, Xilnex, or any inventory CSV / Excel export (`Item.csv`).
- **E-Commerce Marketplaces**: Lazada Price & Stock reports (`pricestock*.xlsx`), Shopee Mass Update (`mass_update_sales_info_*.xlsx`), TikTok Shop.
- **Online Stores**: Google Sites, Shopify, WooCommerce, or custom storefronts.

It eliminates common AI pitfalls (hallucinated SKUs, false stock discrepancies, conflating formulation lines) by enforcing **9 strict invariants**.

---

## The 9 Zero-Mistake Guardrails

1. **Strict Brand Isolation**: 1:1 matching for major brands; negative brand protection prevents third-party items from falling into parent brands; word boundary enforcement (`\bON\b`).
2. **Mandatory Sub-Series Disambiguation (CRITICAL)**: Prevents cross-matching formulations sharing a brand name (e.g., `100% Isolate` vs. `Whey Blend` vs. `Casein`; `Standard Mass` vs. `Carnivor Mass` vs. `Big Steer 1250`).
3. **Packaging Unit Isolation**: Single-serving items (`1 Serving`, `Single Serving`, `1 Bar`, `1 Tub`) must NEVER match multi-packs (`Pack of 20`, `Box of 12`).
4. **Size / Weight Exact Match**: Exact weight/volume matching with standard metric/imperial conversions (`2.27KG` = `5LBS`, `2KG` = `4.4LBS`, `907G` = `2LBS`).
5. **Flavor Strict Jaccard Overlap**: Token overlap threshold (Jaccard >= 0.60) disallows matching generic base words when modifiers exist (e.g., `Double Rich Chocolate` vs. `Extreme Milk Chocolate`).
6. **Strict Canonical Grounding & No Phantom SKUs (CRITICAL)**: Every item code reported in an audit must exist in the canonical POS dataset. Unstocked variations marked `Out of Stock` on website templates are recognized as not carried, not stock discrepancies.
7. **Sub-Series & Formulation Cross-Contamination Protection**: Formulations under the same brand umbrella never borrow flavours or stock levels from each other.
8. **Web Scrape & Line-Break Text Resiliency**: Pre-processes multi-line split prices (e.g., `RM1\n99\n.00` or `RM\n338.00`) and headings before regex parsing.
9. **Mandatory Fresh-State Verification for "Remaining" Lists**: Forbids carrying over stale checklist items across iterations. Re-asserts every remaining item in code against fresh live crawl data.

---

## File Structure

```
stock-reconciliation/
├── SKILL.md                               # The agent skill definition and instruction manual
├── README.md                              # Documentation and usage guide
├── .gitignore                             # Ignores pycache and sensitive spreadsheet files
├── resources/                             # Deterministic lookup tables & schema rules
│   ├── servings_weight_equivalence.json   # Supplement serving-to-weight mapping
│   ├── shopee_template_spec.json          # Marketplace header & column specifications
│   └── subline_guardrails.json            # Distinctive subline tokens & isolation rules
└── scripts/
    ├── reconcile.py                       # Python reconciliation engine (SKU-aware, multi-channel)
    ├── reconcile_shopee_pos.ps1           # Native high-speed Windows reconciliation script
    ├── generate_shopee_upload.ps1         # Upload-ready template generator
    ├── ReconcilerEngine.cs                # Compiled C# matching engine (zero-dependency)
    └── ReconcilerEngine.dll               # Precompiled high-performance binary
```

---

## Running the Reconciliation

### Option 1: Native Windows / PowerShell (Zero Dependencies)
Requires no Python or package installation. Reconciles thousands of rows in milliseconds:
```powershell
powershell -ExecutionPolicy Bypass -File scripts/reconcile_shopee_pos.ps1 -ShopeeMassUpdatePath "path/to/mass_update.xlsx" -PosCatalogPath "path/to/Item.csv"
```
*To generate upload-ready Excel files, add the `-ExportFiles` switch.*

### Option 2: Python Engine
```bash
# Run with python (requires openpyxl):
python scripts/reconcile.py "path/to/Item.csv" "path/to/marketplace_export.xlsx"

# Or via uv:
uv run --with openpyxl python scripts/reconcile.py "path/to/Item.csv" "path/to/marketplace_export.xlsx"
```
