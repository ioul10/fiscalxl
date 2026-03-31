"""FiscalXL — Convertisseur PDF fiscal → Excel (extraction directe)"""
import streamlit as st
import tempfile, os
from core.pdf_to_excel import convert
from utils.logger import get_logger

logger = get_logger(__name__)

st.set_page_config(page_title="FiscalXL", page_icon="📊", layout="wide")

st.markdown("""<style>
.hdr{background:linear-gradient(135deg,#1F3864,#2E75B6);padding:1.4rem 2rem;
  border-radius:12px;margin-bottom:1.2rem;}
.hdr h1{color:white;margin:0;font-size:1.8rem;}
.hdr p{color:#BDD7EE;margin:.3rem 0 0;font-size:.88rem;}
.kpi{background:white;border:1px solid #BDD7EE;border-radius:8px;padding:.7rem;text-align:center;}
.kpi .v{font-size:1.1rem;font-weight:bold;color:#1F3864;}
.kpi .l{font-size:.72rem;color:#888;margin-top:.2rem;}
.ok{background:#E2EFDA;border:1px solid #70AD47;border-radius:8px;padding:.9rem 1.3rem;color:#375623;}
.er{background:#FCE4D6;border:1px solid #C55A11;border-radius:8px;padding:.9rem 1.3rem;color:#7B2C00;}
div[data-testid="stDownloadButton"] button{
  background:linear-gradient(135deg,#1F3864,#2E75B6);color:white;
  border:none;padding:.8rem 2.5rem;font-size:1rem;border-radius:8px;width:100%;}
</style>""", unsafe_allow_html=True)

st.markdown("""<div class="hdr">
<h1>📊 FiscalXL</h1>
<p>Convertisseur PDF → Excel · Pièces annexes IS — MCN loi 9-88 Maroc</p>
</div>""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ℹ️ Comment ça marche")
    st.markdown("""
**Extraction directe par coordonnées**
- Lit chaque mot et sa position dans le PDF
- Reconstitue les lignes automatiquement
- Détecte Actif / Passif / CPC sans config
- Génère un Excel propre et formaté

**Formats supportés**
- Tous formats MCN loi 9-88
- DGI, AMMC, cabinets comptables
    """)
    st.caption("FiscalXL · MCN loi 9-88")

col1, col2 = st.columns([3, 2])
with col1:
    st.markdown("### 📂 Importer le PDF")
    uploaded = st.file_uploader("Glissez-déposez ou cliquez", type=["pdf"])
with col2:
    st.markdown("### 🔄 Pipeline")
    for step, desc in [
        ("1 · Lecture",  "Extraction des mots + positions bbox"),
        ("2 · Lignes",   "Groupement par coordonnée Y"),
        ("3 · Colonnes", "Séparation label / valeurs par X"),
        ("4 · Excel",    "4 feuilles formatées professionnellement"),
    ]:
        st.markdown(
            f'<div style="background:#f8f9fa;border-left:4px solid #1F3864;'
            f'padding:.5rem .8rem;border-radius:0 6px 6px 0;margin:.3rem 0;">'
            f'<strong style="color:#1F3864">{step}</strong><br>'
            f'<span style="color:#555;font-size:.83rem">{desc}</span></div>',
            unsafe_allow_html=True)

if not uploaded:
    st.markdown("""<div style="text-align:center;padding:3rem;color:#888;
      border:2px dashed #BDD7EE;border-radius:12px;background:#f8fafd;margin-top:1rem;">
      <div style="font-size:3rem;">📄</div>
      <h3 style="color:#2E75B6;">Importez un PDF pour commencer</h3>
      <p>Fonctionne avec tous les formats MCN — DGI, AMMC, cabinets comptables</p>
    </div>""", unsafe_allow_html=True)
    st.stop()

st.markdown("---")

with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
    tmp.write(uploaded.getbuffer())
    pdf_path = tmp.name
output_path = pdf_path.replace(".pdf", "_out.xlsx")
excel_bytes = None

try:
    progress = st.progress(0)
    status   = st.empty()

    status.info("📄 Extraction en cours...")
    progress.progress(30)

    with st.spinner("Conversion..."):
        stats = convert(pdf_path, output_path)

    progress.progress(90)

    with open(output_path, "rb") as f:
        excel_bytes = f.read()

    progress.progress(100)
    status.empty()

    info     = stats['info']
    raison   = info.get('raison_sociale', '—')[:30]
    exercice = info.get('exercice_fin', '—')

    for col, (lbl, val) in zip(st.columns(4), [
        ("Raison Sociale",     raison),
        ("Fin exercice",       exercice),
        ("Sections extraites", str(stats['tables'])),
        ("Lignes",             str(stats['rows'])),
    ]):
        col.markdown(f'<div class="kpi"><div class="v">{val}</div>'
                     f'<div class="l">{lbl}</div></div>', unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(f"""<div class="ok">
      ✅ <strong>Conversion réussie !</strong>
      &nbsp;·&nbsp; {stats['tables']} sections (Actif / Passif / CPC)
      &nbsp;·&nbsp; {stats['rows']} lignes extraites
      &nbsp;·&nbsp; {stats['pages']} pages traitées
    </div>""", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("### ⬇️ Télécharger")
    raison_slug = raison.replace(" ", "_")[:20]
    date_slug   = exercice.replace("/", "-")
    fname       = f"{raison_slug}_{date_slug}_fiscal.xlsx"

    st.download_button(
        "📥 Télécharger le fichier Excel",
        data=excel_bytes, file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

except Exception as e:
    logger.exception("Erreur")
    st.markdown(f'<div class="er">❌ <strong>Erreur :</strong> <code>{e}</code></div>',
                unsafe_allow_html=True)
    import traceback; st.code(traceback.format_exc())
finally:
    for f in [pdf_path, output_path]:
        try:
            if os.path.exists(f): os.unlink(f)
        except: pass
