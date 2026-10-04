---
name: branch-restock
description: >-
  Reconcile physical branch POS inventory exports (e.g. Item.csv or stock count exports)
  with the ProteinLab central warehouse / website catalog (proteinlab.com.my) to generate
  verified restock recommendation lists (AV0 for out-of-stock items, AV1 for low-stock items with 1 unit left).
  Enforces strict sub-line disambiguation and live stock thresholds.
---

# Branch Inventory Restock Reconciliation Skill

This skill governs the automated process of auditing physical retail branch inventory against the central warehouse catalog ([proteinlab.com.my](https://www.proteinlab.com.my/)) to produce accurate, discrepancy-free restock orders.

---

## 1. Stock Status Classifications & Thresholds

When evaluating products for branch restocking, strictly apply these thresholds:

| Classification | Branch POS Stock (`SOHQ`) | Central Warehouse Stock (Website) | Default Request Qty | Purpose |
| :--- | :--- | :--- | :--- | :--- |
| **`AV0`** (Out of Stock / < 1) | **0.0000** (or missing from active POS, but present in branch master catalog) | **$> 4$** ($\ge 5$ units) | **2 pcs** (or 5–10 for single-serve sachets/bars/chips) | Replenish completely depleted catalog items. |
| **`AV1`** (Low Stock = 1) | Exactly **1.0000** | **$> 4$** ($\ge 5$ units) | **1 pc** | Bring branch buffer back up to safe operating floor of 2 units. |
| **`Sufficient`** | $\ge 2.0000$ | Any | **0 pcs** | **EXCLUDE**. Never recommend restocking items with $\ge 2$ units on hand. |

> [!IMPORTANT]
> **Warehouse Threshold ($> 4$)**: Never recommend restocking an item if the central website has $\le 4$ units available. This reserves buffer stock for direct online fulfillment.

---

## 2. Automatic Website Connection & Live Stock Probing

**You do NOT need to export or upload any website data.**
The skill automatically connects to **proteinlab.com.my** directly over the web:
- Each product is linked to its unique website `VariantId` stored in [restock_catalog_map.json](./resources/restock_catalog_map.json).
- The script queries the live warehouse cart endpoint in real-time (`POST https://www.proteinlab.com.my/cart/add` with `id=<variant_id>&quantity=1000`).
- The response returns the live, up-to-the-second warehouse stock.
- The live results are cached locally to ensure lightning-fast execution, and can be probed live on demand with `-LiveWebCheck`.

---

## 3. Mandatory Disambiguation & Guardrails

To prevent false restock alerts caused by fuzzy matching between sister product lines:

### A. Strict Sub-Line Distinctiveness
- Tokens such as `Breathe`, `Energy`, `Sprint`, `Hydro`, `Plant`, `Clear`, `Hardcore`, `Plus`, `Iso`, `Mass`, `Decanate`, etc., are **mandatory differentiating keys**, NEVER optional modifiers or fuzzy stop-words.
- **Example**: `Applied Nutrition Endurance Breathe (Cola)` and `Applied Nutrition Endurance Energy (Cola)` are two distinct physical products. **NEVER** bind `Breathe` to `Energy`, even though they share brand, product type (gel), and flavor (Cola).

### B. Multi-Candidate POS Disambiguation
- When an e-commerce or website listing maps to multiple potential POS rows (same Brand + Flavor):
  1. Priority 1: Match by exact **Barcode**.
  2. Priority 2: Match by exact **Item Code / SKU**.
  3. Priority 3: Match exact sub-line tokens. If POS candidate has a conflicting sub-line token that the website product lacks (or vice versa), reject the candidate.
  4. If no exact sub-line match exists, leave the item as unlinked rather than forcing a fuzzy match to a sister product line.

### C. Pre-Alert Stock Check Guardrail
- Before outputting an `AV0` or `AV1` recommendation, verify all POS items in that brand and flavor family.
- If the branch POS already possesses an exact match with healthy stock ($SOHQ \ge 2$), **NEVER** emit a restock alert based on a misaligned sister variant.

### D. Verbatim Ground-Truth Execution (Anti-Hallucination)
- Always pipe row-by-row output directly from verified PowerShell scripts (`reconcile_restock.ps1` or generated report files).
- **NEVER** manually type or reconstruct item rows, codes, or stock quantities from memory. Every reported row and quantity must be directly read from the data source.

### E. Promo & Clearance Exclusion Guardrail
- Automatically exclude all products designated as Promo Clearance (e.g. `*Promo Clearance*`, `*Promo Clearence*`, `(Clearance)`, expiry clearance, hardened stock) from restock recommendations.
- Physical retail branches replenish standard, regular inventory unless clearance restock is explicitly requested with `-IncludeClearance`.

### F. Central Catalog Completeness & Live Probe Guardrail
- The central mapping (`restock_catalog_map.json`) must maintain active coverage of all core stocked brands (Rule 1, Optimum Nutrition, Scitec, MuscleRulz, Core Champs, MMX, etc.).
- When auditing `AV0`, verify that recent warehouse restocks are captured by running live probes (`-LiveWebCheck`) against the central website rather than trusting stale local cache files.
- Never assume an item is permanently unavailable or low stock at the warehouse: shipments arrive regularly, requiring live verification.

### G. Multi-Size URL Invariant (Separate Product Pages per Size)
- Supplement e-commerce stores (such as EasyStore/Shopify on proteinlab.com.my) frequently create distinct product URLs for different weight tiers (e.g., `1lbs`, `2lbs`, `4lbs`, `5lbs`, `10lbs`, and `single-serving sachets`).
- **NEVER** assume a brand's 5lbs listing represents all sizes. When auditing a brand family, sweep all packaging sizes across the store's collection (`/collections/<brand>`) to guarantee smaller bags (e.g. 1lbs Whey Blend) and larger tubs are evaluated.

---

## 4. Standard Reconciliation Workflow (Exhaustive Row-by-Row Execution)

To guarantee that **zero rows are missed**:

### Pass 1: Row-by-Row Sweep of Item.csv (AV1 Low-Stock Audit)
1. Ingests all rows in `Item.csv` (all 400+ rows).
2. Filters every row where `SOHQ == 1.0000` (typically ~149 rows).
3. Evaluates each row individually:
   - Matches against the central website catalog by Barcode, Item Code, or strict Sub-Line & Flavor tokens.
   - Evaluates central warehouse stock:
     * If Warehouse Stock $> 4$: Recommended as **`AV1`** (Request: 1 pc).
     * If Warehouse Stock $\le 4$: Excluded with logged reason (insufficient warehouse stock).
     * If Not Listed Online: Excluded with logged reason (not carried online).
4. Logs exact audit metrics: `Total Rows Scanned`, `SOHQ=1 Rows Audited`, `AV1 Qualified`.

### Pass 2: Row-by-Row Sweep of Central Website Catalog (AV0 Out-of-Stock Audit)
1. Ingests all variants on **proteinlab.com.my** (360+ variants).
2. Filters variants where Central Warehouse Stock is $> 4$.
3. Evaluates each variant individually against branch records:
   - Checks if branch POS has $SOHQ \ge 2$: **Sufficient Stock $\rightarrow$ Excluded** (with pre-alert sub-line safety check).
   - Checks if branch POS has $SOHQ == 1$: Captured under **AV1**.
   - Checks if branch POS has $SOHQ == 0$ or exists in the branch master catalog (`all_stock_count_list.xml`): Recommended as **`AV0`** (Request: 2 pcs, or 5/10 for singles).
4. Logs exact audit metrics: `Master Catalog Items Checked`, `AV0 Qualified`.
```powershell
# Fast run using cached live stock:
powershell -ExecutionPolicy Bypass -File "C:\Users\user\.gemini\config\skills\branch-restock\scripts\reconcile_restock.ps1" -Mode BOTH

# Real-time run querying proteinlab.com.my live over the web:
powershell -ExecutionPolicy Bypass -File "C:\Users\user\.gemini\config\skills\branch-restock\scripts\reconcile_restock.ps1" -LiveWebCheck -Mode BOTH
```
Parameters:
- `-Mode`: `AV0`, `AV1`, or `BOTH`.
- `-LiveWebCheck`: Hits the live website in real-time.
- `-IncludeClearance`: Switch to include Promo Clearance lines (defaults to false / excluded).
- `-ItemCsvPath`: Path to the branch POS file (defaults to `C:\Users\user\Downloads\Item.csv`).

### Step 3: Format & Deliver Output
Check user prompt for formatting preference:
- **In-Chat Summary**: (e.g. *"no need generate excel for me write the summary at here enough"*):
  - Deliver clean Markdown grouped by Brand:
    `Product Name (Website) - Flavor/Option x<Qty>pc AV<0|1> (Main Stock: X)`
- **Excel Spreadsheet**:
  - If requested or generating files, export via Excel COM / PowerShell with columns:
    `Product Name (Website)`, `Flavour / Option`, `Unit Price (RM)`, `Branch Stock`, `Main Branch Stock (Website)`, `Request Qty`, `Est. Total (RM)`, `Branch Master Description`.
