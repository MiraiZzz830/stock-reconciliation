# Retail & E-Commerce Inventory Skills

> **Production-grade AI agent skills for deterministic Point-of-Sale (POS) catalog matching, marketplace stock synchronization, and live branch replenishment orders with 0 mistakes.**

[![Skills](https://img.shields.io/badge/Antigravity-Skills-blue.svg)](https://github.com/MiraiZzz830/stock-reconciliation)
[![Status](https://img.shields.io/badge/Status-Production--Ready-green.svg)](https://github.com/MiraiZzz830/stock-reconciliation)

---

## Included Skills

This repository provides two specialized, standalone inventory skills:

| Skill | Directory | Purpose |
| :--- | :--- | :--- |
| **`stock-reconciliation`** | Root & `skills/stock-reconciliation` | Reconciles POS exports (`Item.csv`) with online marketplace exports (**Shopee**, **Lazada**, **TikTok Shop**). Prevents overselling with SKU priority matching, ghost stock elimination, and duplicate listing tracking. |
| **`branch-restock`** | `branch-restock/` & `skills/branch-restock` | Audits physical retail branch inventory against central warehouse stock ([proteinlab.com.my](https://www.proteinlab.com.my/)) to generate verified **AV0** (out of stock) and **AV1** (low stock) restock orders. |

---

## Repository Structure

```
stock-reconciliation/
├── README.md                              # Repository overview and documentation
├── .gitignore                             # Ignores pycache and sensitive spreadsheets
├── SKILL.md                               # Root skill definition for stock-reconciliation
├── resources/                             # Marketplace & unit conversion specifications
│   ├── servings_weight_equivalence.json   # Supplement serving-to-weight lookup table
│   ├── shopee_template_spec.json          # Marketplace header & column specifications
│   └── subline_guardrails.json            # Distinctive subline tokens & isolation rules
├── scripts/                               # Marketplace reconciliation engines
│   ├── reconcile_shopee_pos.ps1           # Native high-speed Windows reconciliation script
│   ├── generate_shopee_upload.ps1         # Upload-ready template generator
│   ├── ReconcilerEngine.cs                # Compiled C# matching engine (zero-dependency)
│   ├── ReconcilerEngine.dll               # Precompiled high-performance binary
│   └── reconcile.py                       # Python reconciliation engine (cross-platform)
│
├── branch-restock/                        # Standalone Branch Restock Skill
│   ├── SKILL.md                           # Branch restock skill definition & guardrails
│   ├── resources/
│   │   └── restock_catalog_map.json       # Central website variant mappings (proteinlab.com.my)
│   └── scripts/
│       └── reconcile_restock.ps1          # Automated branch restock engine (AV0 / AV1 orders)
│
└── skills/                                # Multi-skill packaging directory
    ├── stock-reconciliation/              # Packaged stock reconciliation skill
    └── branch-restock/                    # Packaged branch restock skill
```

---

## 1. Using `stock-reconciliation`

Reconciles physical POS exports against Shopee or Lazada marketplace files.

### Native Windows / PowerShell (Recommended, Zero Dependencies)
```powershell
powershell -ExecutionPolicy Bypass -File scripts/reconcile_shopee_pos.ps1 -ShopeeMassUpdatePath "path/to/mass_update.xlsx" -PosCatalogPath "path/to/Item.csv"
```
*To generate upload-ready Excel files, add `-ExportFiles`.*

### Python Engine
```bash
python scripts/reconcile.py "path/to/Item.csv" "path/to/marketplace_export.xlsx"
```

---

## 2. Using `branch-restock`

Audits retail branch POS stock against live central warehouse inventory ([proteinlab.com.my](https://www.proteinlab.com.my/)).

```powershell
# Fast run using cached live stock:
powershell -ExecutionPolicy Bypass -File branch-restock/scripts/reconcile_restock.ps1 -Mode BOTH

# Real-time run querying proteinlab.com.my live over the web:
powershell -ExecutionPolicy Bypass -File branch-restock/scripts/reconcile_restock.ps1 -LiveWebCheck -Mode BOTH

# Audit low stock only (AV1 = 1 pc left at branch):
powershell -ExecutionPolicy Bypass -File branch-restock/scripts/reconcile_restock.ps1 -Mode AV1

# Audit depleted stock only (AV0 = 0 pcs left at branch):
powershell -ExecutionPolicy Bypass -File branch-restock/scripts/reconcile_restock.ps1 -Mode AV0
```

---

## Core Zero-Mistake Guardrails

1. **Strict Brand Isolation**: 1:1 matching for major brands; negative brand protection prevents third-party items from falling into parent brands.
2. **Mandatory Sub-Series Disambiguation**: Enforces distinct formulation boundaries (`Sprint` vs. `Energy` vs. `Breathe`, `Isolate` vs. `Blend` vs. `Casein`).
3. **Packaging Unit Isolation**: Single-serving items (`1 Serving`, `1 Bar`) never match multi-packs (`Pack of 20`).
4. **Ghost Stock Elimination**: Automatically zeros out marketplace stock for POS items with $0$ on hand.
5. **Multi-Listing Clone Safeguard**: Aggregates duplicate allocations and calculates total overselling risk.
6. **Pre-Alert Stock Verification**: Never emits restock alerts if the branch POS already possesses an exact match with healthy stock ($SOHQ \ge 2$).
7. **Zero Stale Memory Carryover**: Programmatically asserts $\text{Remaining Missing List} \cap \text{Freshly Matched Items} = \emptyset$ on iterative audits.
8. **Summary-First Rule**: Outputs structured tables directly in chat by default; generates spreadsheet files on-demand only.
