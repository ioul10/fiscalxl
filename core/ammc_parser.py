"""
core/ammc_parser.py
Parseur format AMMC — PDF avec extract_tables().
Structure : 7 cols actif, 5 cols passif, 7-8 cols CPC.
Lettres décoratives verticales en col 0, headers sur 3-4 lignes.
"""

import re
import unicodedata
import pdfplumber
from utils.logger import get_logger

logger = get_logger(__name__)

# ── Détection sections ────────────────────────────────────────────────────────
ACTIF_KW  = ["bilan actif modele", "bilan (actif", "bilan actif",
              "tableau n  1 1 2", "tableau n 1 1 2",
              "immobilisations en non valeurs", "actif immobilise"]
PASSIF_KW = ["bilan passif modele", "bilan (passif", "bilan passif",
              "tableau n  1 2 2", "tableau n 1 2 2",
              "capitaux propres", "passif circulant"]
CPC_KW    = ["compte de produits", "tableau n  2", "tableau n 2",
              "produits d exploitation", "charges d exploitation"]

# ── Skip ──────────────────────────────────────────────────────────────────────
HEADER_SKIP = {
    "a c t i f", "p a s s i f", "exercice", "exercice precedent",
    "brut", "net", "amortissements et provisions",
    "designation", "operations", "totaux de l exercice",
    "propres a l exercice", "concernant les exercices precedents",
    "1", "2", "3 = 2 + 1", "4",
}
SKIP_PREFIX = (
    "tableau n", "bilan (", "compte de produits",
    "agence du", "identifiant fiscal", "exercice du",
    "1)variation", "2)achats",
)

TOTAL_KW = ("total i", "total ii", "total iii", "total general",
            "total iv", "total v ", "total vi", "total vii",
            "total viii", "total ix", "total des produits",
            "total des charges", "total (a+b")
RESULT_KW = ("resultat d exploitation", "resultat financier",
             "resultat courant", "resultat non courant",
             "resultat avant impot", "resultat net",
             "impots sur les")
SECTION_KW = ("produits d exploitation", "charges d exploitation",
              "produits financiers", "charges financieres",
              "produits non courant", "charges non courant",
              "immobilisations en non", "immobilisations incorporelles",
              "immobilisations corporelles", "immobilisations financiere",
              "ecarts de conversion", "stocks", "creances de l actif",
              "titres valeurs", "tresorerie", "capitaux propres",
              "dettes de financement", "dettes du passif")


# ── Utilitaires ───────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    s = unicodedata.normalize('NFD', s)
    s = s.encode('ascii', 'ignore').decode('utf-8').lower()
    s = re.sub(r'[^\w\s]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def _is_rotated(v) -> bool:
    """Détecte les lettres décoratives verticales ex: 'A\nC\nT\nI\nF'."""
    if not v:
        return False
    parts = [p.strip() for p in str(v).split('\n') if p.strip()]
    return len(parts) >= 3 and all(len(p) <= 2 and p.isalpha() for p in parts)

def _parse_num(s) -> float | None:
    if s is None:
        return None
    s = str(s).strip().replace('\xa0', '').replace('\n', '')
    if not s or s in ['-', '—', '']:
        return None
    neg = s.startswith('-')
    s = s.lstrip('-').lstrip('+')
    s = s.replace(' ', '')
    if re.match(r'^\d{1,3}(\.\d{3})*,\d{2}$', s):
        s = s.replace('.', '').replace(',', '.')
    elif re.match(r'^\d+,\d{2}$', s):
        s = s.replace(',', '.')
    elif re.match(r'^\d+(\.\d+)?$', s):
        pass
    else:
        return None
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None

def _clean_label(s) -> str:
    if not s:
        return ''
    s = str(s).replace('\n', ' ').strip()
    s = re.sub(r'\s{2,}', ' ', s)
    # Enlever numéros romains en début (I, II, III...)
    s = re.sub(r'^(I{1,3}|IV|V|VI{1,3}|IX|X{1,3})\s+', '', s)
    return s.strip()

def _should_skip(label: str) -> bool:
    n = _norm(label)
    if not n or len(n) < 2:
        return True
    if n in HEADER_SKIP:
        return True
    if any(n.startswith(p) for p in SKIP_PREFIX):
        return True
    if re.match(r'^[\d\s\(\)\+\=\/\-]+$', n):
        return True
    if len(n.replace(' ', '')) <= 2:
        return True
    return False

def _row_type(label: str) -> str:
    n = _norm(label)
    if any(n.startswith(k) for k in TOTAL_KW) or 'total general' in n:
        return 'total'
    if any(n.startswith(k) for k in RESULT_KW):
        return 'result'
    if any(n.startswith(k) for k in SECTION_KW):
        return 'section'
    stripped = label.strip()
    if stripped and stripped == stripped.upper() and len(stripped) > 4:
        return 'section'
    return 'normal'


# ── Parseurs par section ──────────────────────────────────────────────────────

def _parse_actif(table: list) -> list:
    """
    Actif AMMC : 7 cols
    [déco, label, brut, None, amort, net_n, net_n1]
    """
    rows = []
    seen = set()
    for row in table:
        if not row or len(row) < 3:
            continue
        # Ignorer lettres décoratives col 0
        if _is_rotated(row[0]):
            continue
        label = _clean_label(row[1] if len(row) > 1 else row[0])
        if not label or _should_skip(label):
            continue
        key = _norm(label)
        if key in seen:
            continue
        seen.add(key)

        n = len(row)
        if n >= 7:
            brut   = _parse_num(row[2])
            amort  = _parse_num(row[4])
            net_n  = _parse_num(row[5])
            net_n1 = _parse_num(row[6])
        elif n >= 5:
            brut   = _parse_num(row[2])
            amort  = _parse_num(row[3])
            net_n  = _parse_num(row[4])
            net_n1 = None
        else:
            continue

        rows.append({
            'label': label,
            'brut': brut, 'amort': amort,
            'net_n': net_n, 'net_n1': net_n1,
            'type': _row_type(label),
        })
    return rows


def _parse_passif(table: list) -> list:
    """
    Passif AMMC : 5 cols
    [déco, label, None, val_n, val_n1]
    """
    rows = []
    seen = set()
    for row in table:
        if not row or len(row) < 3:
            continue
        if _is_rotated(row[0]):
            continue
        label = _clean_label(row[1] if len(row) > 1 else row[0])
        if not label or _should_skip(label):
            continue
        key = _norm(label)
        if key in seen:
            continue
        seen.add(key)

        n = len(row)
        if n >= 5:
            val_n  = _parse_num(row[3])
            val_n1 = _parse_num(row[4])
        elif n >= 4:
            val_n  = _parse_num(row[2])
            val_n1 = _parse_num(row[3])
        else:
            continue

        rows.append({
            'label': label,
            'val_n': val_n, 'val_n1': val_n1,
            'type': _row_type(label),
        })
    return rows


def _parse_cpc(tables: list) -> list:
    """
    CPC AMMC : pages 4 et 5, 7-8 cols
    [déco, num_romain, label, propre_n, (vide), prec_n, total_n, total_n1]
    """
    rows = []
    seen = set()

    for table in tables:
        for row in table:
            if not row or len(row) < 4:
                continue
            if _is_rotated(row[0]):
                continue

            n = len(row)
            # Détecter position du label selon nb de colonnes
            if n >= 7:
                label = _clean_label(row[2])
                if n >= 8:
                    propre_n = _parse_num(row[3])
                    prec_n   = _parse_num(row[5])
                    total_n  = _parse_num(row[6])
                    total_n1 = _parse_num(row[7])
                else:
                    propre_n = _parse_num(row[3])
                    prec_n   = _parse_num(row[4])
                    total_n  = _parse_num(row[5])
                    total_n1 = _parse_num(row[6])
            elif n >= 5:
                label    = _clean_label(row[1])
                propre_n = _parse_num(row[2])
                prec_n   = _parse_num(row[3])
                total_n  = _parse_num(row[4])
                total_n1 = None
            else:
                continue

            if not label or _should_skip(label):
                continue
            key = _norm(label)
            if key in seen:
                continue
            seen.add(key)

            rows.append({
                'label': label,
                'propre_n': propre_n, 'prec_n': prec_n,
                'total_n': total_n, 'total_n1': total_n1,
                'type': _row_type(label),
            })

    return rows


# ── Extraction info ───────────────────────────────────────────────────────────

def _extract_info(pdf) -> dict:
    info = {}
    for i in range(min(2, len(pdf.pages))):
        text = pdf.pages[i].extract_text() or ''
        if not info.get('raison_sociale'):
            m = re.search(r'[Rr]aison\s+[Ss]ociale\s*[:\-]?\s*([A-ZÀ-Ü][^\n]{3,60})', text)
            if m:
                info['raison_sociale'] = m.group(1).strip()
        if not info.get('identifiant_fiscal'):
            m = re.search(r'[Ii]dentifiant\s+[Ff]iscal\s*[:\-]?\s*(\d+)', text)
            if m:
                info['identifiant_fiscal'] = m.group(1)
        if not info.get('taxe_professionnelle'):
            m = re.search(r'[Tt]axe\s+[Pp]rof\w*\.?\s*[:\-]?\s*([\d\s]+)', text)
            if m:
                info['taxe_professionnelle'] = m.group(1).strip().replace(' ', '')
        if not info.get('adresse'):
            m = re.search(r'[Aa]dresse\s*[:\-]?\s*([^\n]{5,60})', text)
            if m:
                info['adresse'] = m.group(1).strip()
        if not info.get('exercice'):
            m = re.search(r'(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})', text)
            if m:
                info['exercice']       = f"Du {m.group(1)} au {m.group(2)}"
                info['exercice_debut'] = m.group(1)
                info['exercice_fin']   = m.group(2)
        # Depuis tableau page 1
        tables = pdf.pages[i].extract_tables()
        for t in tables:
            for row in t:
                cells = [str(c).strip() if c else '' for c in row]
                j = ' '.join(cells).lower()
                if 'raison sociale' in j and not info.get('raison_sociale'):
                    vals = [c for c in cells if c and 'raison' not in c.lower() and ':' not in c and len(c) > 3]
                    if vals:
                        info['raison_sociale'] = vals[-1]
                elif 'identifiant fiscal' in j and not info.get('identifiant_fiscal'):
                    vals = [c for c in cells if re.match(r'^\d{5,}$', c)]
                    if vals:
                        info['identifiant_fiscal'] = vals[0]
                elif 'taxe professionnelle' in j and not info.get('taxe_professionnelle'):
                    vals = [c for c in cells if re.match(r'^\d{5,}$', c)]
                    if vals:
                        info['taxe_professionnelle'] = vals[0]
                elif 'adresse' in j and not info.get('adresse'):
                    vals = [c for c in cells if c and 'adresse' not in c.lower() and ':' not in c and len(c) > 5]
                    if vals:
                        info['adresse'] = vals[-1]

    info.setdefault('raison_sociale', '')
    info.setdefault('identifiant_fiscal', '')
    info.setdefault('taxe_professionnelle', '')
    info.setdefault('adresse', '')
    info.setdefault('exercice', '')
    info.setdefault('exercice_fin', '')
    return info


# ── Point d'entrée ────────────────────────────────────────────────────────────

def parse(pdf_path: str) -> dict:
    """
    Parse un PDF format AMMC.
    Retourne : {info, actif, passif, cpc, pages}
    """
    pdf = pdfplumber.open(pdf_path)
    n   = len(pdf.pages)

    info = _extract_info(pdf)

    # Détecter les pages par section
    actif_tables  = []
    passif_tables = []
    cpc_tables    = []

    for i, page in enumerate(pdf.pages):
        text = _norm(page.extract_text() or '')
        tables = page.extract_tables()
        if any(k in text for k in ACTIF_KW):
            actif_tables.extend(tables)
        elif any(k in text for k in PASSIF_KW):
            passif_tables.extend(tables)
        elif any(k in text for k in CPC_KW):
            cpc_tables.extend(tables)

    actif  = _parse_actif(actif_tables[0])  if actif_tables  else []
    passif = _parse_passif(passif_tables[0]) if passif_tables else []
    cpc    = _parse_cpc(cpc_tables)

    pdf.close()

    logger.info(f"AMMC parsed: {len(actif)} actif, {len(passif)} passif, {len(cpc)} cpc")
    return {
        'info': info,
        'actif': actif,
        'passif': passif,
        'cpc': cpc,
        'pages': n,
        'format': 'AMMC',
    }
