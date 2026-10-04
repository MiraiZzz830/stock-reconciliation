# Stock Reconciliation Skill

> **Deterministic inventory and catalog reconciliation between POS exports and online channels (Marketplaces & Web Stores) with 0 mistakes.**

[![Skill](https://img.shields.io/badge/Antigravity-Skill-blue.svg)](https://github.com/MiraiZzz830/stock-reconciliation)
[![Status](https://img.shields.io/badge/Status-Production--Ready-green.svg)](https://github.com/MiraiZzz830/stock-reconciliation)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://python.org)

---

## Overview

The **Stock Reconciliation Skill** provides deterministic matching and discrepancy auditing between:
- **POS Systems**: SQL POS, Autocount, Xilnex, or any inventory CSV / Excel export.
- **E-Commerce Marketplaces**: Lazada Price & Stock reports, Shopee Mass Update / Sales Info, TikTok Shop.
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
├── SKILL.md             # The agent skill definition and instruction manual
├── scripts/
│   └── reconcile.py     # Deterministic reconciliation engine
├── .gitignore           # Ignores pycache and sensitive spreadsheet files
└── README.md            # Documentation and usage guide
```

---

## Quick Usage

### Prerequisites
- Python 3.10+
- `openpyxl`

### Running the Reconciliation Script
```bash
# Run via uv:
uv run --with openpyxl python scripts/reconcile.py "<path_to_pos_file.csv>" "<path_to_marketplace_file.xlsx>"

# Or with standard python:
python scripts/reconcile.py "path/to/pos_export.csv" "path/to/marketplace_export.xlsx"
```

---

## License

Private repository for internal use.
