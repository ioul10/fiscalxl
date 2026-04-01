"""
core/dgi_parser.py
Parseur format DGI — PDF avec tableaux fusionnés / cellules multi-lignes.
Utilise extract_words() + coordonnées bbox.
"""

import re
import unicodedata
from collections import defaultdict

import pdfplumber
from utils.logger import get_logger

logger = get_logger(__name__)

# ── Mots-clés détection sections ─────────────────────────────────────────────
ACTIF_KW  = ["actif immobilise", "immobilisations en non valeur",
              "bilan (actif", "bilan actif", "immobilisation incorporelle",
              "b i l a n (actif", "tableau n  01 (1", "tableau n 01 (1",
              "frais preliminaires"]
PASSIF_KW = ["capitaux propres", "bilan (passif", "bilan passif",
              "b i l a n (passif", "tableau n  01 (2", "tableau n 01 (2"]
CPC_KW    = ["produits d exploitation", "charges d exploitation",
             "compte de produits", "ventes de marchandises",
             "tableau n  02", "tableau n 02"]

# ── Skip ──────────────────────────────────────────────────────────────────────
SKIP_EXACT = {
    "brut", "net", "designation", "operations",
    "a c t i f", "p a s s i f",
    "propres a l exercice", "concernant les exercices precedents",
    "totaux de l exercice", "totaux de l exercice precedent",
    "amortissements et provisions", "exercice precedent",
    "1", "2", "3 = 2 + 1", "4", "3 = 1 + 2",
}
SKIP_PREFIX = (
    "tableau n", "bilan (", "compte de produits",
    "identifiant", "exercice du", "1)variation", "2)achat",
    "cadre reserve", "signature", "nb :",
)
SKIP_SUFFIX = ("(1/2)", "(2/2)", "(hors taxes)", "(suite)")

TOTAL_KW = ("total i ", "total ii", "total iii", "total general",
            "total (a+b", "total iv", "total v ",
            "total vi", "total vii", "total viii", "total ix",
            "total des produits", "total des charges")
RESULT_KW = ("resultat d exploitation", "resultat financier",
             "resultat courant", "resultat non courant",
             "resultat avant impot", "resultat net",
             "impots sur les")
SECTION_KW = ("produits d exploitation", "charges d exploitation",
              "produits financiers", "charges financieres",
              "produits non courant", "charges non courant",
              "immobilisations en non", "immobilisation incorporelle",
              "immobilisations corporelles", "immobilisations financiere",
              "ecarts de conversion", "stocks", "creances de l actif",
              "titres et valeurs", "titres valeurs", "tresorerie",
              "capitaux propres", "dettes de financement",
              "dettes du passif")


# ── Utilitaires ───────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    s = unicodedata.normalize('NFD', s)
    s = s.encode('ascii', 'ignore').decode('utf-8').lower()
    s = re.sub(r'[^\w\s]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def _parse_num_tokens(tokens: list) -> float | None:
    if not tokens:
        return None
    s = ''.join(tokens).replace(' ', '').replace('\xa0', '')
    if not s or s in ['-', '—']:
        return None
    neg = (s.startswith('(') and s.endswith(')')) or s.startswith('-')
    if s.startswith('(') and s.endswith(')'):
        s = s[1:-1]
    if s.startswith('-'):
        s = s[1:]
    s = s.replace(',', '.')
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None

def _should_skip(label: str) -> bool:
    n = _norm(label)
    if not n or len(n) < 2:
        return True
    if n in SKIP_EXACT:
        return True
    if any(n.startswith(p) for p in SKIP_PREFIX):
        return True
    if any(n.endswith(s) for s in SKIP_SUFFIX):
        return True
    if re.match(r'^[\d\s\(\)\+\-\=\/]+$', n):
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

def _clean_label(label: str) -> str:
    label = label.strip()
    label = re.sub(r'^[A-Z]{1,2}\s+(?=[A-ZÀ-Ü])', '', label)
    label = re.sub(r'^(I{1,3}|IV|V|VI{1,3}|IX|X{1,3})\s+', '', label)
    return label.strip()


# ── Extraction par bbox ───────────────────────────────────────────────────────

class _WordExtractor:
    def __init__(self, page, label_ratio=0.40):
        self.pw    = page.width
        self.lmax  = self.pw * label_ratio
        self._words = page.extract_words(x_tolerance=3, y_tolerance=3)

    def _lines(self, y_tol=3):
        groups = defaultdict(list)
        for w in self._words:
            groups[round(w['top'] / y_tol) * y_tol].append(w)
        return {y: sorted(ws, key=lambda w: w['x0']) for y, ws in groups.items()}

    def _tokens_in(self, ws, x0, x1):
        return [w['text'] for w in ws if x0 <= w['x0'] < x1]

    def extract(self, mode: str, skip=8) -> list:
        lines  = self._lines()
        bounds = self._col_bounds(mode)
        rows, seen = [], set()

        for i, y in enumerate(sorted(lines)):
            if i < skip:
                continue
            ws = lines[y]
            label_tokens = [w['text'] for w in ws if w['x0'] < self.lmax]
            label = _clean_label(' '.join(label_tokens))
            if not label or _should_skip(label):
                continue
            key = _norm(label)
            if key in seen:
                continue
            seen.add(key)

            vals = [_parse_num_tokens(self._tokens_in(ws, x0, x1))
                    for x0, x1 in bounds]
            row = {'label': label, 'type': _row_type(label)}
            if mode == 'actif':
                row.update({'brut': vals[0], 'amort': vals[1],
                             'net_n': vals[2], 'net_n1': vals[3]})
            elif mode == 'passif':
                row.update({'val_n': vals[0], 'val_n1': vals[1]})
            elif mode == 'cpc':
                row.update({'propre_n': vals[0], 'prec_n': vals[1],
                             'total_n': vals[2], 'total_n1': vals[3]})
            rows.append(row)
        return rows

    def _col_bounds(self, mode: str) -> list:
        pw = self.pw
        if mode == 'actif':
            return [(pw*.440, pw*.560), (pw*.560, pw*.680),
                    (pw*.680, pw*.790), (pw*.790, pw*.940)]
        elif mode == 'passif':
            return [(pw*.620, pw*.780), (pw*.780, pw*.940)]
        elif mode == 'cpc':
            return [(pw*.430, pw*.560), (pw*.560, pw*.690),
                    (pw*.690, pw*.810), (pw*.810, pw*.940)]
        return []


# ── Extraction info ───────────────────────────────────────────────────────────

def _extract_info(pdf) -> dict:
    info = {}
    for i in range(min(3, len(pdf.pages))):
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
    Parse un PDF format DGI.
    Retourne : {info, actif, passif, cpc, pages}
    """
    pdf = pdfplumber.open(pdf_path)
    n   = len(pdf.pages)
    info = _extract_info(pdf)

    actif_pages, passif_pages, cpc_pages = [], [], []
    for i, page in enumerate(pdf.pages):
        t = _norm(page.extract_text() or '')
        if any(k in t for k in ACTIF_KW):
            actif_pages.append(i)
        if any(k in t for k in PASSIF_KW):
            passif_pages.append(i)
        if any(k in t for k in CPC_KW):
            cpc_pages.append(i)

    # Fallbacks
    if not actif_pages:
        actif_pages = list(range(1, min(3, n)))
    if not passif_pages:
        passif_pages = list(range(2, min(5, n)))
    if not cpc_pages:
        cpc_pages = list(range(3, min(7, n)))

    def _extract(pages, mode):
        rows, seen = [], set()
        for idx in pages:
            if idx >= n:
                continue
            for row in _WordExtractor(pdf.pages[idx]).extract(mode):
                key = _norm(row['label'])
                if key not in seen:
                    seen.add(key)
                    rows.append(row)
        return rows

    actif  = _extract(actif_pages,  'actif')
    passif = _extract(passif_pages, 'passif')
    cpc    = _extract(cpc_pages,    'cpc')

    pdf.close()
    logger.info(f"DGI parsed: {len(actif)} actif, {len(passif)} passif, {len(cpc)} cpc")
    return {
        'info': info,
        'actif': actif,
        'passif': passif,
        'cpc': cpc,
        'pages': n,
        'format': 'DGI',
    }
