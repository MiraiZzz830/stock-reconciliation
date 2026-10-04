---
name: stock-reconciliation
description: >-
  Deterministically checks, matches, and reconciles stock between POS system exports (Item.csv/Excel) and online channels (Lazada, Shopee, TikTok Shop, Google Sites, Shopify, custom web stores) with 0 mistakes. Enforces strict brand matching, zero phantom SKUs, canonical POS grounding, mandatory sub-series differentiation (e.g., Sprint vs. Energy vs. Breathe, Isolate vs. Blend vs. Mass Gainer), packaging unit isolation, zero stale memory carryover on iterative audits, and outputs discrepancies directly in-chat without unnecessary Excel files.
---

# Stock Reconciliation Skill (Zero-Mistake Protocol)

Use this skill whenever the user asks to check, match, or reconcile inventory and catalog offerings between a **POS system** (e.g. `Item.csv` or POS Excel exports) and an **online sales channel** (e-commerce marketplaces like Lazada, Shopee, TikTok Shop, or web stores like Google Sites, Shopify, WooCommerce).

---

## 1. Zero-Mistake Guardrails (Mandatory Invariants)

When matching items between POS and online channels, NEVER rely on visual guessing, loose keyword searches, or assumptions about product ranges. You MUST enforce the following 10 invariants:

### Guardrail 1: Strict Brand Isolation
* Known brands (e.g., `Applied Nutrition`, `Optimum Nutrition`, `Rule 1`, `Scitec Nutrition`, `Core Champs`, `MuscleRulz`, `MuscleMeds`, `Proscience`, `SuppzLab`, `SIS`, `BSN`, `GAT Sport`, `MMX`, etc.) must match 1:1.
* **Negative Brand Protection**: If an online listing mentions an external third-party brand (e.g., `Cellucor`, `Redcon1`, `JNX The Curse`, `Scivation Xtend`, `Mutant`, `OstroVit`), it **must NEVER** fall back into a known parent brand.
* **Word Boundary Enforcement**: Match brands using whole words or specific URL slugs (e.g., `\bON\b` or slug `optimum-nutrition`), never naive substring inclusion.

### Guardrail 2: Mandatory Sub-Series Disambiguation (CRITICAL)
Products sharing a brand, category, and flavor often belong to completely different chemical formulations or sub-lines. They must **never** be cross-matched:
* **Energy Gels**: `Sprint` vs. `Energy` vs. `Breathe` (e.g., `Endurance Sprint Orange` CANNOT match `Endurance Energy Orange`).
* **Proteins**: `100% Isolate` vs. `Whey Blend / Concentrate` vs. `Hydrolyzed / Hydro` vs. `Casein` vs. `Clear Whey`.
* **Mass Gainers**: `Standard Mass` vs. `Serious Mass` vs. `Big Steer 1250` vs. `Carnivor Mass`.

### Guardrail 3: Packaging Unit Isolation
* Single-serving items (`1 Serving`, `Single Serving`, `1 Bar`, `1 Bag`, `1 Tub`, `1 Sachet`) must **NEVER** be matched with multi-packs (`Pack of 20`, `Pack of 12`, `Pack of 6`, `Box of 15`).
* If an online listing is for a multi-pack and POS only tracks single units, mark the multi-pack as `Unmatched / Pack Multiplier Required`.

### Guardrail 4: Size, Weight & Servings Equivalence
* Convert metric/imperial equivalents if standard: `2.27KG` = `5LBS`, `2.6KG` = `5.6LBS / 6LBS`, `2KG` = `4.4LBS`, `907G` = `2LBS`.
* **Servings-to-Weight Domain Equivalence** ([servings_weight_equivalence.json](./resources/servings_weight_equivalence.json)):
  - `36 Servings` $\leftrightarrow$ `1.1KG` (Keep Moment Daily Whey)
  - `14 Servings` $\leftrightarrow$ `450G` (Critical Whey 450g)
  - `25 Servings` $\leftrightarrow$ `825G` (Critical Whey 825g)
  - `35 Servings` $\leftrightarrow$ `875G` (Clear Whey 875g)
  - `50 Servings` $\leftrightarrow$ `250G` (Creatine / Powder)
  - `56 Servings` $\leftrightarrow$ `4.5LBS` / `2KG` (MuscleMeds Carnivor)
  - `60 Servings` $\leftrightarrow$ `300G` / `2.4LBS`
  - `67 Servings` $\leftrightarrow$ `2KG` (Critical Whey 2kg)
  - `74 Servings` $\leftrightarrow$ `5LBS` (ON Gold Standard Whey)
  - `76 Servings` $\leftrightarrow$ `5LBS` (Rule 1 R1 Protein)
  - `100 Servings` $\leftrightarrow$ `7LBS` / `7.3LBS` / `3.3KG` (MuscleMeds Carnivor 7lbs)

### Guardrail 5: Flavor & Attribute Strict Jaccard
* Flavors must share high token overlap (Jaccard >= 0.60).
* Disallow matching purely on generic words like "Chocolate" or "Vanilla" when distinguishing modifiers exist (e.g., `Double Rich Chocolate` vs. `Extreme Milk Chocolate` vs. `Chocolate Peanut Butter`).

### Guardrail 6: Strict Canonical Grounding & No Phantom SKUs (CRITICAL)
* **Zero Hallucinated Codes**: Every item code/SKU reported in an audit or discrepancy table MUST strictly exist in the loaded canonical POS records (`item['Code'] in pos_codes`). Never invent, extrapolate, or synthesize a POS code from a website variation name.
* **Asymmetrical Absence ≠ Stock Discrepancy**: If an online listing or website page template shows a flavor or size as `(Out of Stock)`, but that variation **does not exist anywhere in the POS database**, it is **NOT** a stock discrepancy. The store simply does not carry that variation. Do not report it as "POS has stock" or demand flipping the website to "In Stock".

### Guardrail 7: Ghost Stock & Discontinued POS Item Elimination
* Any marketplace listing variation that has positive stock ($> 0$), but whose POS item has been discontinued, removed, or has $SOHQ = 0$:
  - Must be explicitly flagged as **Ghost Stock**.
  - Its stock must be updated to **0** in marketplace update templates to prevent overselling.

### Guardrail 8: Duplicate Listing / Clone Safeguard
* Sellers frequently clone products across multiple Product IDs pointing to the same physical POS SKU.
* Track total marketplace allocation: if $\sum \text{Marketplace Stock} > \text{POS SOHQ}$, emit a **High Risk Overselling Warning**.

### Guardrail 9: Web Scrape & Line-Break Text Resiliency
* Web builders (Google Sites, Webflow, Shopify themes) frequently introduce line-break splits inside price strings (e.g., `RM1\n99\n.00` or `RM\n338.00`) and variation headings. Parsers must normalize multi-line whitespace before extracting prices or flavor blocks to avoid false price discrepancies.

### Guardrail 10: Mandatory Fresh-State Verification for "Remaining" Lists (Zero Stale Memory Carryover)
* **Zero Memory Residue**: When presenting follow-up audit results after a re-crawl or newly published update, **NEVER** copy or carry forward items from previous conversation turns into the "What Remains / Missing" list.
* **Programmatic Assertion Before Output**: Every item asserted as "Missing" or "Discrepant" must be re-checked against the active, freshly parsed dataset in code. If the live page already lists the variant (e.g. `Weight: 600g (120 Servings)`, `RM238.00`, `(In Stock)`), it must be confirmed as matched and immediately dropped from the missing checklist.
* **Audit Sanity Rule**:
  $$\text{Remaining Missing List} \cap \text{Freshly Matched Items} = \emptyset$$

---

## 2. Execution Workflow

When running a catalog or stock reconciliation:
1. Locate or ask for:
   - POS file (e.g. `Item.csv` or `.user_uploaded/*.csv`)
   - Comparison channel: Marketplace file (`pricestock*.xlsx`, `mass_update_sales_info_*.xlsx`) OR Online Store URL / crawled pages JSON.
2. Load POS records deterministically (preserves all SKUs, leading zeros, and row counts).
3. If reconciling against a Web Store / Google Sites:
   - Pre-process and normalize text lines to heal split prices (`RM` + `digits`).
   - Group pages strictly by brand and sub-series slug.
   - Ground every potential discrepancy strictly against the canonical POS list. If a flavor is marked "Out of Stock" on the web and does not exist in POS, filter it out as unstocked catalog template text.
4. If reconciling against a Marketplace (Lazada / Shopee / TikTok):
   - **Fast Native Windows Engine (Zero Dependencies, Recommended)**:
     ```powershell
     powershell -ExecutionPolicy Bypass -File "$env:USERPROFILE\.gemini\config\skills\stock-reconciliation\scripts\reconcile_shopee_pos.ps1" -ShopeeMassUpdatePath "<path_to_marketplace_excel>" -PosCatalogPath "<path_to_Item.csv>"
     ```
   - **Python Engine**:
     ```bash
     python "$env:USERPROFILE\.gemini\config\skills\stock-reconciliation\scripts\reconcile.py" "<path_to_pos_csv>" "<path_to_marketplace_excel>"
     ```
5. Filter and categorize results:
   - **In Sync**: `Online Stock == POS SOHQ` and `Price == POS Price 1`.
   - **Real Discrepancies**: Online price != POS Price 1, or genuine POS item with SOH > 0 marked "Out of Stock".
   - **Unstocked Online Variations**: Listed online but not carried in POS (safe to remain Out of Stock or delete).
   - **POS Only**: Physical stock in POS not yet listed on the online channel.
6. **Iterative Re-Check Sanity Gate (Post-Publish)**:
   - Whenever the user republishes or refreshes the site, re-run full extraction on the new crawl.
   - Re-evaluate all checklist items against the active dataset.
   - Automatically promote fixed items to the "Successfully Resolved" section and purge them from the "Remaining" section.

---

## 3. In-Chat Output Standards

* **Default to In-Chat Tables**: Write the discrepancy list directly in the chat response organized by Brand.
* **Do NOT generate Excel files by default**: Only generate an Excel report or ready-to-upload template if the user explicitly asks for an Excel file.
* **Format Prices Correctly**: Use `RM` and avoid unescaped `$` signs to protect KaTeX rendering.
* **Table Columns**:
  | Product Name | Variation | POS Code | Online Price / Stock | POS Price / Stock | Discrepancy Note |
