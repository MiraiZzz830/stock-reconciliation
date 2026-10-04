import openpyxl
import csv
import re
import sys
import os

ALL_BRANDS = [
    ("OPTIMUM NUTRITION", ["OPTIMUM NUTRITION", "ON ", "(ON)"]),
    ("SCITEC NUTRITION", ["SCITEC NUTRITION", "SCITEC"]),
    ("APPLIED NUTRITION", ["APPLIED NUTRITION"]),
    ("CORE CHAMPS", ["CORE CHAMPS"]),
    ("MUSCLERULZ", ["MUSCLERULZ", "MUSCLE RULZ"]),
    ("MUSCLEMEDS", ["MUSCLEMEDS", "CARNIVOR"]),
    ("RULE 1", ["RULE 1", "RULE ONE", "R1"]),
    ("PROSCIENCE", ["PROSCIENCE NUTRA", "PROSCIENCE"]),
    ("SUPPZLAB", ["SUPPZLAB", "PROTEINLAB", "PROTEIN LAB", "TEAM PROTEIN LAB", "IRON WILL", "IRONWILL"]),
    ("BASIC SUPPLEMENTS", ["BASIC SUPPLEMENTS"]),
    ("FIT & LEAN", ["FIT & LEAN", "FIT AND LEAN"]),
    ("SIS", ["SIS", "SCIENCE IN SPORT"]),
    ("QUEST NUTRITION", ["QUEST NUTRITION", "QUEST"]),
    ("LEGENDARY", ["LEGENDARY"]),
    ("OPI", ["OPI"]),
    ("DYMATIZE", ["DYMATIZE", "ISO 100", "ISO100"]),
    ("MMX", ["MMX", "MUSCLE METABOLIX"]),
    ("UNIVERSAL", ["UNIVERSAL", "ANIMAL"]),
    ("NUTREX", ["NUTREX RESEARCH", "NUTREX", "LIPO 6", "LIPO-6"]),
    ("BSN", ["BSN"]),
    ("EVERBUILD", ["EVERBUILD NUTRITION", "EVERBUILD"]),
    ("APOTEC", ["APOTEC"]),
    ("DEXTER JACKSON", ["DEXTER JACKSON"]),
    ("CYCLONE CUP", ["CYCLONECUP", "CYCLONE CUP"]),
    ("BPI SPORTS", ["BPI SPORTS", "BPI"]),
    ("PERFECT SHAKER", ["PERFECT SHAKER"]),
    ("PURE SCIENCE", ["PURE SCIENCE LABS", "PURE SCIENCE"]),
    ("ZENDAVA", ["ZENDAVA"]),
    ("MUSCLETECH", ["MUSCLETECH", "NITRO TECH", "NITRO-TECH"]),
    ("KEEP MOMENT", ["KEEP MOMENT"]),
    ("BE STRONG", ["BE STRONG"]),
    ("DREAM TAN", ["DREAM TAN"]),
    ("FIT FIRST", ["FIT FIRST"]),
    ("GAT SPORT", ["GAT SPORT", "GAT"]),
    ("KEVIN LEVRONE", ["KEVIN LEVRONE"]),
    ("MHP", ["MHP"]),
    ("FULFIL", ["FULFIL NUTRITION", "FULFIL"]),
    ("LEAN BODY", ["LEAN BODY", "LABRADA"]),
    ("ULTIMATE NUTRITION", ["ULTIMATE NUTRITION"]),
    ("XNERGY", ["XNERGY"]),
    ("CELLUCOR", ["CELLUCOR", "C4"]),
    ("REDCON1", ["REDCON1", "TOTAL WAR"]),
    ("JNX", ["JNX", "THE CURSE"]),
    ("SCIVATION", ["SCIVATION", "XTEND"]),
    ("MAXLER", ["MAXLER"]),
    ("PURUS LABS", ["PURUS LABS", "PURUS"]),
    ("SAN NATION", ["SAN NATION", "SAN"]),
    ("OSTROVIT", ["OSTROVIT"]),
    ("BAAM", ["BAAM"]),
    ("PROSUPPS", ["PROSUPPS", "HYDE"]),
    ("KAGED", ["KAGED"]),
    ("MUTANT", ["MUTANT"]),
    ("ETERNAL WELLNESS", ["ETERNAL WELLNESS"]),
]

SIZES = [
    ("16LBS", ["16LBS", "16LB"]),
    ("15LBS", ["15LBS", "15LB"]),
    ("12LBS", ["12LBS", "12LB"]),
    ("11.46LBS", ["11.46LBS", "11.46LB"]),
    ("10LBS", ["10LBS", "10LB"]),
    ("7LBS", ["7LBS", "7.4LBS", "7.4LB", "7.3LBS", "3.3KG", "100 SERVINGS", "100SERVINGS", "100 SV", "100SV"]),
    ("6LBS", ["6LBS", "5.73LBS", "5.73LB", "6LB"]),
    ("5LBS", ["5LBS", "5.02LBS", "5.03LBS", "5.09LBS", "5.01LBS", "5.28LBS", "5.2LBS", "4.95LBS", "4.99LBS", "5.02LB", "5.03LB", "2.27KG", "2.35KG", "2350G", "2350GM", "2270G", "2KG", "67 SERVINGS", "67SERVINGS", "67 SV", "67SV", "74 SERVINGS", "74SERVINGS", "76 SERVINGS", "76SERVINGS"]),
    ("4LBS", ["4LBS", "4.4LBS", "4.4LB", "4.5LBS", "4.5LB", "4.19LBS", "4.09LBS", "3.97LBS", "3.96LBS", "3.8LBS", "3.7LBS", "3.61LBS", "3.52LBS", "3.5LBS", "4LB", "56 SERVINGS", "56SERVINGS", "56 SV", "56SV"]),
    ("3LBS", ["3LBS", "3.06KG", "3060G", "3520G", "3520GM", "1.36KG"]),
    ("2LBS", ["2LBS", "1.98LBS", "1.95LBS", "2.01LBS", "2.4LBS", "1.87LBS", "1.91LBS", "1.98LB", "1.95LB", "2.01LB", "907G", "1KG", "1.1KG", "1025G", "920G", "875G", "825G", "2LB", "36 SERVINGS", "36SERVINGS", "36 SV", "36SV", "35 SERVINGS", "35SERVINGS", "25 SERVINGS", "25SERVINGS"]),
    ("1LBS", ["1LBS", "1.43LBS", "1.48LBS", "0.99LBS", "1.06LBS", "1.76LBS", "0.99LB", "1.48LB", "450G", "500G", "600G", "1LB", "14 SERVINGS", "14SERVINGS", "14 SV", "14SV"]),
    ("500G", ["500G", "50 SERVINGS", "50SERVINGS", "50 SV", "50SV", "250G", "300G", "60 SERVINGS", "60SERVINGS"]),
    ("SINGLE SERVING", ["SINGLE SERVING", "1 BAG", "1 TUB", "1S", "60ML", "220ML", "500ML", "1 PACK", "1 SACHET", "1 CAN", "1 BAR"]),
    ("1L", ["1L", "1000ML"]),
    ("750ML", ["750ML"]),
    ("700ML", ["700ML", "739ML"]),
    ("600ML", ["600ML"]),
    ("500ML", ["500ML"]),
    ("400ML", ["400ML", "470ML"]),
    ("SIZE 2XL", ["SIZE 2XL", "2XL", "XXL"]),
    ("SIZE XL", ["SIZE XL", "XL"]),
    ("SIZE L", ["SIZE L", "L"]),
    ("SIZE M", ["SIZE M", "M"]),
    ("SIZE S", ["SIZE S", "S"]),
    ("SIZE XS", ["SIZE XS", "XS"]),
    ("FREE SIZE", ["FREE SIZE"])
]

PRODUCT_TYPES = [
    ("SKIPPING ROPE", ["SKIPPING ROPE", "ROPE"]),
    ("PLASTIC BAG", ["PLASTIC BAG"]),
    ("BAG", ["DRY BAG", "BAG"]),
    ("SHAKER", ["SHAKER", "SHAKER BOTTLE", "BOTTLE"]),
    ("BELT", ["BELT"]),
    ("STRAP", ["LIFTING STRAPS", "STRAPS", "STRAP"]),
    ("WRAP", ["WRIST WRAP", "WRAP"]),
    ("GLOVE", ["GLOVE", "GLOVES"]),
    ("TOWEL", ["TOWEL"]),
    ("YOGA MAT", ["YOGA MAT", "MAT"]),
    ("T-SHIRT", ["T-SHIRT", "TSHIRT", "TEE", "SHIRT"]),
    ("SINGLET", ["SINGLET"]),
    ("JACKET", ["JACKET", "SWEATER", "WINDBREAKER"]),
    ("PROTEIN BAR", ["PROTEIN BAR", "BAR", "CRISP"]),
    ("PROTEIN CHIPS", ["PROTEIN CHIPS", "CHIPS"]),
    ("PROTEIN DONUTS", ["DONUTS", "DONUT"]),
    ("ENERGY GEL", ["ENERGY GEL", "GEL", "ISOTONIC"]),
    ("ENERGY DRINK", ["ENERGY DRINK", "DRINK", "CAN", "SPARKLING", "WATER", "READY TO DRINK"]),
    ("MASS GAINER", ["MASS GAINER", "MASS BUILD", "MASS", "BIG STEER", "WEIGHT GAINER", "PRO GAINER", "GAIN RULZ"]),
    ("ISOLATE", ["ISOLATE", "ISO BUILD", "ISO RULZ", "ISO 100", "HYDROWHEY", "ANABOLIC ISO", "HYDRO ISOLATE", "CLEAR WHEY", "ISO WHEY"]),
    ("CASEIN", ["CASEIN"]),
    ("WHEY PROTEIN", ["WHEY PROTEIN", "WHEY", "CRITICAL WHEY", "NITRA WHEY", "HYBRID PRO", "100% WHEY"]),
    ("PLANT PROTEIN", ["PLANT PROTEIN", "VEGAN PROTEIN", "PLANT"]),
    ("CREATINE", ["CREATINE", "CREAPURE"]),
    ("COLLAGEN", ["COLLAGEN"]),
    ("GLUTAMINE", ["GLUTAMINE"]),
    ("CARNITINE", ["L-CARNITINE", "CARNITINE"]),
    ("PRE WORKOUT", ["PRE WORKOUT", "PRE-WORKOUT", "PUMP", "PRE LIFT", "PRE RULZ", "HOT BLOOD", "RDX"]),
    ("AMINO / BCAA / EAA", ["AMINO", "BCAA", "EAA"]),
    ("CARBOTEIN", ["CARBOTEIN", "GLYCOGEN"]),
    ("SARMS", ["SARMS", "MK-677", "MK677"]),
    ("TESTO", ["TESTO", "TESTO PUNCH", "HOS TRIO"]),
    ("VITAMIN / MULTI", ["MULTIPRO", "MULTI PRO", "ELECTROLYTE"])
]

def extract_brand_strict(text):
    text_upper = " " + text.upper() + " "
    for brand_canon, aliases in ALL_BRANDS:
        for alias in aliases:
            pattern = r'(?<![A-Z0-9])' + re.escape(alias.upper().strip()) + r'(?![A-Z0-9])'
            if re.search(pattern, text_upper):
                return brand_canon
    return "UNKNOWN"

def extract_size(text):
    text_upper = " " + text.upper().replace('(', ' ').replace(')', ' ') + " "
    for size_canon, aliases in SIZES:
        for alias in aliases:
            pattern = r'(?<![A-Z0-9])' + re.escape(alias.upper().strip()) + r'(?![A-Z0-9])'
            if re.search(pattern, text_upper):
                return size_canon
    return "UNKNOWN"

def extract_product_type(text):
    text_upper = " " + text.upper() + " "
    for ptype, aliases in PRODUCT_TYPES:
        for alias in aliases:
            pattern = r'(?<![A-Z0-9])' + re.escape(alias.upper().strip()) + r'(?![A-Z0-9])'
            if re.search(pattern, text_upper):
                return ptype
    return "OTHER"

def get_specific_subtype(text):
    t = text.upper()
    if 'SPRINT' in t: return 'SPRINT'
    if 'BREATHE' in t: return 'BREATHE'
    if 'ISOLATE' in t or 'ISO BUILD' in t or 'ISO RULZ' in t or 'ISO 100' in t: return 'ISOLATE'
    if 'HYDRO' in t: return 'HYDRO'
    if 'CASEIN' in t: return 'CASEIN'
    if 'CRITICAL' in t: return 'CRITICAL'
    if 'CLEAR' in t: return 'CLEAR'
    if 'SYNTHA' in t: return 'SYNTHA'
    if 'SERIOUS MASS' in t: return 'SERIOUS MASS'
    if 'TRUE MASS' in t: return 'TRUE MASS'
    if 'BIG STEER' in t: return 'BIG STEER'
    return None

def is_multipack(text):
    return bool(re.search(r'PACK OF \d+|BOX OF \d+', text.upper()))

def clean_flavor_string(s):
    if not s:
        return set()
    s = s.upper()
    s = s.replace('RICHCHOCOLATE', 'RICH CHOCOLATE')
    s = s.replace('SRAWBERRY', 'STRAWBERRY')
    s = s.replace('STROM', 'STORM')
    s = s.replace('MIN CHOCOLATE', 'MINT CHOCOLATE')
    s = s.replace('CHOCO', 'CHOCOLATE')
    s = s.replace('CHOC', 'CHOCOLATE')
    s = s.replace('STRAW', 'STRAWBERRY')
    s = s.replace('VANI', 'VANILLA')
    s = s.replace('CARAM', 'CARAMEL')
    s = s.replace('P/BUTTER', 'PEANUT BUTTER')
    s = s.replace('P/B', 'PEANUT BUTTER')
    s = s.replace('PB', 'PEANUT BUTTER')
    s = s.replace('&', ' AND ')
    s = s.replace('+', ' PLUS ')
    s = s.replace('/', ' ')
    s = s.replace('N CREAM', 'AND CREAM')
    s = s.replace('N CREME', 'AND CREAM')
    s = s.replace('CRÈME', 'CREAM')
    s = s.replace('CREME', 'CREAM')
    s = s.replace('UNFLAVORED', 'UNFLAVORED')
    s = s.replace('UNFLAVOURED', 'UNFLAVORED')
    s = re.sub(r'[^A-Z0-9]+', ' ', s)
    tokens = set(s.split())
    noise = {'SHAKE', 'POWDER', 'FLAVOR', 'FLAVOURED', 'FLAVORED', 'PROTEIN', 'ICE', 'STYLE', 'SERIES'}
    return tokens - noise

def reconcile_inventory(pos_file_path, marketplace_file_path):
    pos_items = []
    pos_by_code = {}
    pos_by_barcode = {}

    with open(pos_file_path, mode='r', encoding='utf-8-sig', errors='replace') as f:
        for row in csv.DictReader(f):
            code = row.get('Code', '').strip()
            barcode = row.get('Barcode', '').strip()
            desc = row.get('Description', '').strip()
            brand = extract_brand_strict(desc)
            size = extract_size(desc)
            ptype = extract_product_type(desc)
            subtype = get_specific_subtype(desc)
            multipack = is_multipack(desc)
            groups = re.findall(r'\((.*?)\)', desc)
            flavor_tokens = set()
            for g in groups:
                if 'SERVING' not in g.upper() and extract_size(g) == 'UNKNOWN' and 'PACK' not in g.upper():
                    flavor_tokens.update(clean_flavor_string(g))

            pos_obj = {
                'code': code,
                'barcode': barcode,
                'desc': desc,
                'sohq': float(row.get('SOHQ', 0) or 0),
                'price1': float(row.get('Price 1', 0) or 0),
                'brand': brand,
                'size': size,
                'ptype': ptype,
                'subtype': subtype,
                'multipack': multipack,
                'flavor_tokens': flavor_tokens,
            }
            pos_items.append(pos_obj)
            if code:
                pos_by_code[code.upper()] = pos_obj
            if barcode:
                pos_by_barcode[barcode.upper()] = pos_obj

    # Open marketplace workbook
    wb = openpyxl.load_workbook(marketplace_file_path, data_only=True)
    sheet = wb.active
    if 'template' in wb.sheetnames:
        sheet = wb['template']

    # Detect headers
    row_start = 1
    col_name = 3
    col_var = 14
    col_sku = 9
    col_stock = 13
    col_price = None
    
    for c in range(1, sheet.max_column + 1):
        v = str(sheet.cell(row=1, column=c).value or '').lower()
        if 'product name' in v or 'item name' in v: col_name = c
        elif 'variation' in v: col_var = c
        elif 'sku' in v and 'shop' not in v: col_sku = c
        elif 'parklane' in v or 'stock' in v or 'quantity' in v: col_stock = c
        elif 'price' in v: col_price = c

    # Lazada 5-row header structure
    if sheet.cell(row=2, column=1).value in ('Optional', 'Mandatory'):
        row_start = 5
    # Shopee 6-row header structure
    elif sheet.cell(row=3, column=1).value is not None and 'mandatory' in str(sheet.cell(row=3, column=1).value).lower():
        row_start = 7

    mkt_items = []
    for i in range(row_start, sheet.max_row + 1):
        name = str(sheet.cell(row=i, column=col_name).value or '').strip()
        sku = str(sheet.cell(row=i, column=col_sku).value or '').strip()
        stock_val = sheet.cell(row=i, column=col_stock).value
        try: stock = float(stock_val if stock_val is not None else 0)
        except: stock = 0.0
        var = str(sheet.cell(row=i, column=col_var).value or '').strip() if col_var else ''
        
        price = 0.0
        if col_price:
            pval = sheet.cell(row=i, column=col_price).value
            try: price = float(pval if pval is not None else 0)
            except: price = 0.0

        if not name and not sku:
            continue

        brand = extract_brand_strict(name)
        size = extract_size(name)
        if size == "UNKNOWN" and var: size = extract_size(var)
        ptype = extract_product_type(name + " " + var)
        subtype = get_specific_subtype(name + " " + var)
        multipack = is_multipack(name + " " + var)
        var_flavor_tokens = clean_flavor_string(var)
        if not var_flavor_tokens and not var:
            for g in re.findall(r'\((.*?)\)', name):
                if extract_size(g) == 'UNKNOWN':
                    var_flavor_tokens.update(clean_flavor_string(g))
                    
        mkt_items.append({
            'row': i,
            'name': name,
            'sku': sku,
            'var': var,
            'stock': stock,
            'price': price,
            'brand': brand,
            'size': size,
            'ptype': ptype,
            'subtype': subtype,
            'multipack': multipack,
            'flavor_tokens': var_flavor_tokens,
        })

    matches = []
    unmatched_mkt = []
    matched_pos_codes = set()

    for m in mkt_items:
        # Tier 1: Deterministic Exact SKU / Barcode Match
        tier1_pos = None
        if m['sku']:
            sku_clean = m['sku'].upper().strip()
            if sku_clean in pos_by_code:
                tier1_pos = pos_by_code[sku_clean]
            elif sku_clean in pos_by_barcode:
                tier1_pos = pos_by_barcode[sku_clean]

        if tier1_pos:
            matches.append((m, tier1_pos, 1.0, "Tier 1: SKU Match"))
            matched_pos_codes.add(tier1_pos['code'])
            continue

        # Tier 2: Deterministic Semantic Match
        best_p = None
        best_score = 0
        for p in pos_items:
            # 1. Exact Brand match
            if m['brand'] == "UNKNOWN" or p['brand'] == "UNKNOWN" or m['brand'] != p['brand']: continue
            # 2. Product Type match
            if m['ptype'] != "OTHER" and p['ptype'] != "OTHER" and m['ptype'] != p['ptype']:
                if not (m['ptype'] in ("ISOLATE", "WHEY PROTEIN") and p['ptype'] in ("ISOLATE", "WHEY PROTEIN")): continue
            # 3. Size match
            if m['size'] != "UNKNOWN" and p['size'] != "UNKNOWN" and m['size'] != p['size']: continue
            # 4. Packaging match (Pack of 20 vs Single Serving)
            if m['multipack'] != p['multipack']: continue
            # 5. Sub-type match (Sprint vs Energy vs Breathe)
            if m['subtype'] != p['subtype']: continue
            
            # 6. Flavor match
            m_f = m['flavor_tokens']
            p_f = p['flavor_tokens']
            if m_f and p_f:
                if not (m_f & p_f): continue
                jacc = len(m_f & p_f) / len(m_f | p_f)
                if jacc < 0.60: continue
                score = jacc
            elif not m_f and not p_f:
                score = 1.0
            else:
                continue
                
            if score > best_score:
                best_score = score
                best_p = p
                
        if best_p:
            matches.append((m, best_p, best_score, "Tier 2: Semantic Match"))
            matched_pos_codes.add(best_p['code'])
        else:
            unmatched_mkt.append(m)

    # Classifications
    in_sync = []
    discrepancies = []
    ghost_stock = []
    
    pos_allocation = {}
    for m, p, score, method in matches:
        pos_allocation.setdefault(p['code'], []).append(m)
        if m['stock'] > 0 and p['sohq'] == 0:
            ghost_stock.append((m, p, "Ghost Stock: POS is 0 but Marketplace has stock"))
        elif m['stock'] == p['sohq']:
            in_sync.append((m, p))
        else:
            discrepancies.append((m, p))

    # Duplicate listing check
    duplicate_warnings = []
    for code, m_list in pos_allocation.items():
        if len(m_list) > 1:
            total_mkt_stock = sum(item['stock'] for item in m_list)
            pos_stock = pos_by_code[code]['sohq']
            if total_mkt_stock > pos_stock:
                duplicate_warnings.append({
                    'code': code,
                    'count': len(m_list),
                    'total_mkt_stock': total_mkt_stock,
                    'pos_stock': pos_stock,
                    'delta': total_mkt_stock - pos_stock
                })

    unmatched_pos = [p for p in pos_items if p['code'] not in matched_pos_codes and p['sohq'] > 0]

    return {
        'total_pos': len(pos_items),
        'total_mkt': len(mkt_items),
        'matched_count': len(matches),
        'in_sync': in_sync,
        'discrepancies': discrepancies,
        'ghost_stock': ghost_stock,
        'duplicate_warnings': duplicate_warnings,
        'unmatched_mkt': unmatched_mkt,
        'unmatched_pos': unmatched_pos,
        'matches': matches
    }

def print_summary(results):
    print("=" * 80)
    print(" STOCK RECONCILIATION AUDIT SUMMARY (ZERO-MISTAKE PROTOCOL)")
    print("=" * 80)
    print(f"Total Marketplace Listings Audited : {results['total_mkt']}")
    print(f"Total Active POS Records Evaluated : {results['total_pos']}")
    print(f"Total Successfully Matched         : {results['matched_count']}")
    print(f"  - In-Sync (Identical Stock)      : {len(results['in_sync'])}")
    print(f"  - Stock Discrepancies            : {len(results['discrepancies'])}")
    print(f"  - Ghost Stock (POS 0, Mkt > 0)   : {len(results['ghost_stock'])}")
    print(f"Unmatched Marketplace Listings     : {len(results['unmatched_mkt'])}")
    print(f"Unlinked POS In-Stock Items        : {len(results['unmatched_pos'])}")
    print("-" * 80)

    if results['ghost_stock']:
        print("\n[!] HIGH RISK GHOST STOCK (Update to 0 on Marketplace immediately):")
        print(f"{'POS Code':<15} | {'Marketplace Name':<40} | {'Mkt Stock':<10} | {'POS SOHQ'}")
        print("-" * 80)
        for m, p, note in results['ghost_stock']:
            name_short = (m['name'][:37] + '...') if len(m['name']) > 40 else m['name']
            print(f"{p['code']:<15} | {name_short:<40} | {m['stock']:<10} | {p['sohq']}")

    if results['duplicate_warnings']:
        print("\n[!] MULTI-LISTING OVERSELLING RISKS (Cloned Listings):")
        for dup in results['duplicate_warnings']:
            print(f"  Code: {dup['code']} across {dup['count']} listings. Sum Mkt Stock: {dup['total_mkt_stock']} > POS SOHQ: {dup['pos_stock']} (Risk: +{dup['delta']})")

    if results['discrepancies']:
        print(f"\n[*] STOCK DISCREPANCIES ({len(results['discrepancies'])} items):")
        print(f"{'POS Code':<15} | {'Marketplace Product':<35} | {'Variation':<20} | {'Mkt':<6} | {'POS':<6} | {'Delta'}")
        print("-" * 95)
        for m, p in results['discrepancies']:
            name_short = (m['name'][:32] + '...') if len(m['name']) > 35 else m['name']
            var_short = (m['var'][:18] + '...') if len(m['var']) > 20 else m['var']
            delta = m['stock'] - p['sohq']
            print(f"{p['code']:<15} | {name_short:<35} | {var_short:<20} | {m['stock']:<6.0f} | {p['sohq']:<6.0f} | {delta:+6.0f}")

    print("\n" + "=" * 80)
    print(" AUDIT COMPLETED DETERMINISTICALLY")
    print("=" * 80)

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: python reconcile.py <path_to_pos_file.csv> <path_to_marketplace_file.xlsx>")
        sys.exit(1)

    pos_file = sys.argv[1]
    mkt_file = sys.argv[2]

    if not os.path.exists(pos_file):
        print(f"Error: POS file not found: {pos_file}")
        sys.exit(1)
    if not os.path.exists(mkt_file):
        print(f"Error: Marketplace file not found: {mkt_file}")
        sys.exit(1)

    res = reconcile_inventory(pos_file, mkt_file)
    print_summary(res)
