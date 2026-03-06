"""
nPOI DAS Configurator - Application Streamlit
Génère un fichier Excel d'organisation des entrées RF sur des unités nPOI (1U Rack 19", 8 ports)

Installation :
    pip install streamlit openpyxl

Lancement :
    streamlit run npoi_configurator.py
"""

import streamlit as st
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
import io

# ── Constantes ────────────────────────────────────────────────────────────────

FREQUENCES_ORDRE = ["700/800", "900", "1800", "2100", "2600", "3500"]

# Palette de couleurs par fréquence (fond, texte)
FREQ_COLORS = {
    "700/800": ("1A3A5C", "7EC8E3"),
    "900":     ("1A4A2A", "7ED9A0"),
    "1800":    ("4A1A1A", "E37E7E"),
    "2100":    ("3A1A4A", "C07EE3"),
    "2600":    ("4A3A1A", "E3C07E"),
    "3500":    ("1A3A4A", "7EC0E3"),
}

# Palette de couleurs par opérateur (fond, texte) - 6 opérateurs max
OP_COLORS = [
    ("1A3A5C", "7EC8E3"),
    ("1A4A2A", "7ED9A0"),
    ("4A1A1A", "E37E7E"),
    ("3A1A4A", "C07EE3"),
    ("4A3A1A", "E3C07E"),
    ("1A3A4A", "7EC0E3"),
]

GRIS_LIBRE   = ("2A2A2A", "888888")
ENTETE_FOND  = "0D1B2A"
ENTETE_TEXTE = "C0CFE8"


# ── Logique métier ─────────────────────────────────────────────────────────────

def construire_ports(nb_secteurs, operateurs, frequences, mimo, tri):
    """Génère la liste ordonnée de tous les ports RF."""
    ports = []
    for s in range(1, nb_secteurs + 1):
        for op in operateurs:
            for freq in frequences:
                chaines = ["V", "H"] if mimo else [None]
                for chaine in chaines:
                    ports.append({
                        "secteur":   f"S{s}",
                        "operateur": op,
                        "frequence": freq,
                        "chaine":    chaine,
                    })

    if tri == "Par fréquence":
        ports.sort(key=lambda p: (
            FREQUENCES_ORDRE.index(p["frequence"]),
            p["secteur"],
            p["operateur"],
        ))
    else:  # Par opérateur
        ports.sort(key=lambda p: (
            p["operateur"],
            p["secteur"],
            FREQUENCES_ORDRE.index(p["frequence"]),
        ))

    return ports


def grouper_en_npoi(ports):
    """Découpe la liste de ports en groupes de 8 (1 nPOI = 8 ports)."""
    return [ports[i:i + 8] for i in range(0, len(ports), 8)]


def label_port(port):
    """Retourne le label complet d'un port."""
    if port is None:
        return "Port libre"
    base = f"{port['secteur']}-{port['operateur']}-{port['frequence']}"
    return f"{base}-{port['chaine']}" if port["chaine"] else base


# ── Génération Excel ───────────────────────────────────────────────────────────

def style_cellule(ws, coord, valeur, bg_hex, fg_hex,
                  bold=False, align="left", border=True, taille=11):
    cell = ws[coord]
    cell.value = valeur
    cell.font = Font(
        name="Consolas", size=taille, bold=bold,
        color=fg_hex
    )
    cell.fill = PatternFill("solid", fgColor=bg_hex)
    cell.alignment = Alignment(
        horizontal=align, vertical="center", wrap_text=False
    )
    if border:
        thin = Side(style="thin", color="1E3A5A")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def generer_excel(nb_secteurs, operateurs, frequences, mimo, tri, npois):
    wb = openpyxl.Workbook()

    # ── Feuille 1 : Détail par nPOI ──────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Configuration nPOI"
    ws1.sheet_view.showGridLines = False
    ws1.row_dimensions[1].height = 30
    ws1.row_dimensions[2].height = 20

    # Titre principal
    ws1.merge_cells("A1:G1")
    style_cellule(ws1, "A1",
                  "nPOI DAS — Configuration des entrées RF",
                  ENTETE_FOND, "E0F0FF",
                  bold=True, align="center", taille=14, border=False)

    # Sous-titre paramètres
    ws1.merge_cells("A2:G2")
    params = (f"Secteurs: {nb_secteurs}  |  "
              f"Opérateurs: {', '.join(operateurs)}  |  "
              f"Fréquences: {', '.join(frequences)}  |  "
              f"Mode: {'MIMO 2×2' if mimo else 'SISO'}  |  "
              f"Tri: {tri}  |  "
              f"nPOI nécessaires: {len(npois)}")
    style_cellule(ws1, "A2", params,
                  "0A1525", "5A8AAA",
                  align="center", taille=9, border=False)

    largeurs = [8, 12, 18, 14, 10, 32, 12]
    for i, w in enumerate(largeurs, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    row = 4
    for idx, npoi in enumerate(npois):
        # Titre nPOI
        ws1.merge_cells(f"A{row}:G{row}")
        ws1.row_dimensions[row].height = 22
        style_cellule(ws1, f"A{row}",
                      f"  nPOI {idx + 1}   —   1U Rack 19\"   —   "
                      f"{len(npoi)}/8 ports utilisés",
                      "0D1B2A", "7EC8E3",
                      bold=True, align="left", taille=11)
        row += 1

        # En-têtes colonnes
        entetes = ["Port", "Secteur", "Opérateur", "Fréquence", "Chaîne", "Label complet", "nPOI"]
        ws1.row_dimensions[row].height = 18
        for col, titre in enumerate(entetes, 1):
            style_cellule(ws1, f"{get_column_letter(col)}{row}",
                          titre, "162030", "8AABCC",
                          bold=True, align="center", taille=10)
        row += 1

        # Ports
        for p in range(8):
            port = npoi[p] if p < len(npoi) else None
            ws1.row_dimensions[row].height = 20

            if port is None:
                bg, fg = GRIS_LIBRE
            elif tri == "Par fréquence":
                bg, fg = FREQ_COLORS.get(port["frequence"], GRIS_LIBRE)
            else:
                op_idx = operateurs.index(port["operateur"]) if port["operateur"] in operateurs else 0
                bg, fg = OP_COLORS[op_idx % len(OP_COLORS)]

            valeurs = [
                f"P{p + 1}",
                port["secteur"]   if port else "—",
                port["operateur"] if port else "—",
                f"{port['frequence']} MHz" if port else "—",
                port["chaine"] if port and port["chaine"] else ("SISO" if port else "—"),
                label_port(port),
                f"nPOI {idx + 1}",
            ]
            for col, val in enumerate(valeurs, 1):
                style_cellule(ws1, f"{get_column_letter(col)}{row}",
                              val, bg, fg, align="center" if col == 1 else "left")
            row += 1

        row += 1  # ligne vide entre nPOI

    # ── Feuille 2 : Matrice synthèse ─────────────────────────────────────────
    ws2 = wb.create_sheet("Matrice synthèse")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:I1")
    style_cellule(ws2, "A1", "Matrice de synthèse — Assignation des ports par nPOI",
                  ENTETE_FOND, "E0F0FF", bold=True, align="center", taille=13, border=False)

    ws2.column_dimensions["A"].width = 12
    for c in range(2, 10):
        ws2.column_dimensions[get_column_letter(c)].width = 26

    # En-têtes
    row2 = 3
    ws2.row_dimensions[row2].height = 18
    style_cellule(ws2, f"A{row2}", "nPOI", "162030", "8AABCC", bold=True, align="center")
    for p in range(8):
        style_cellule(ws2, f"{get_column_letter(p + 2)}{row2}",
                      f"Port {p + 1}", "162030", "8AABCC", bold=True, align="center")
    row2 += 1

    for idx, npoi in enumerate(npois):
        ws2.row_dimensions[row2].height = 18
        style_cellule(ws2, f"A{row2}", f"nPOI {idx + 1}",
                      "0D1B2A", "7EC8E3", bold=True, align="center")
        for p in range(8):
            port = npoi[p] if p < len(npoi) else None
            if port is None:
                bg, fg = GRIS_LIBRE
            elif tri == "Par fréquence":
                bg, fg = FREQ_COLORS.get(port["frequence"], GRIS_LIBRE)
            else:
                op_idx = operateurs.index(port["operateur"]) if port["operateur"] in operateurs else 0
                bg, fg = OP_COLORS[op_idx % len(OP_COLORS)]

            style_cellule(ws2, f"{get_column_letter(p + 2)}{row2}",
                          label_port(port), bg, fg, align="center")
        row2 += 1

    # ── Feuille 3 : Légende ──────────────────────────────────────────────────
    ws3 = wb.create_sheet("Légende")
    ws3.sheet_view.showGridLines = False
    ws3.column_dimensions["A"].width = 20
    ws3.column_dimensions["B"].width = 30

    ws3.merge_cells("A1:B1")
    style_cellule(ws3, "A1", "Légende des couleurs",
                  ENTETE_FOND, "E0F0FF", bold=True, align="center", taille=13, border=False)

    if tri == "Par fréquence":
        ws3["A3"] = None
        style_cellule(ws3, "A2", "Couleur", "162030", "8AABCC", bold=True, align="center")
        style_cellule(ws3, "B2", "Fréquence", "162030", "8AABCC", bold=True, align="center")
        for i, freq in enumerate(frequences):
            bg, fg = FREQ_COLORS.get(freq, GRIS_LIBRE)
            style_cellule(ws3, f"A{3 + i}", f"{freq} MHz", bg, fg, align="center")
            style_cellule(ws3, f"B{3 + i}", f"Bande {freq} MHz", bg, fg)
    else:
        style_cellule(ws3, "A2", "Couleur", "162030", "8AABCC", bold=True, align="center")
        style_cellule(ws3, "B2", "Opérateur", "162030", "8AABCC", bold=True, align="center")
        for i, op in enumerate(operateurs):
            bg, fg = OP_COLORS[i % len(OP_COLORS)]
            style_cellule(ws3, f"A{3 + i}", op, bg, fg, align="center")
            style_cellule(ws3, f"B{3 + i}", op, bg, fg)

    # Retour en bytes
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Interface Streamlit ────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="nPOI DAS Configurator",
        page_icon="📡",
        layout="wide",
    )

    # CSS personnalisé
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Share+Tech+Mono&family=Exo+2:wght@400;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Share Tech Mono', monospace;
        background-color: #0a0e1a;
        color: #c0cfe8;
    }
    .stApp { background-color: #0a0e1a; }

    h1, h2, h3 { font-family: 'Exo 2', sans-serif; color: #7ec8e3; }

    .block-container { padding-top: 2rem; }

    div[data-testid="stSidebar"] {
        background-color: #0d1520;
        border-right: 1px solid #1e3a5a;
    }

    .stButton > button {
        background: linear-gradient(135deg, #1e4a7a, #0a2a4a);
        color: #7ec8e3;
        border: 1px solid #2a6aaa;
        border-radius: 4px;
        font-family: 'Share Tech Mono', monospace;
        font-size: 14px;
        padding: 8px 20px;
        width: 100%;
        transition: all 0.2s;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #2a6aaa, #1a4a7a);
        color: #e0f0ff;
    }

    .port-card {
        border-radius: 4px;
        padding: 8px 12px;
        margin: 3px 0;
        font-family: 'Share Tech Mono', monospace;
        font-size: 12px;
        border-left: 3px solid;
    }

    .npoi-header {
        background: #0d1b2a;
        border: 1px solid #1e3a5a;
        border-radius: 6px;
        padding: 10px 16px;
        margin: 16px 0 8px 0;
        color: #7ec8e3;
        font-family: 'Exo 2', sans-serif;
        font-weight: 700;
        font-size: 15px;
    }

    div[data-testid="stMultiSelect"] div[role="option"] { font-size: 13px; }

    .stSelectbox label, .stMultiSelect label, .stSlider label,
    .stRadio label, .stCheckbox label { color: #8aabcc !important; font-size: 13px; }
    </style>
    """, unsafe_allow_html=True)

    # ── Titre ──────────────────────────────────────────────────────────────────
    st.markdown("""
    <div style="background:linear-gradient(135deg,#0d1b2a,#1a2a3a);
                border-bottom:1px solid #1e3a5a;
                padding:20px 28px;margin-bottom:24px;border-radius:8px;">
        <span style="font-family:'Exo 2',sans-serif;font-size:24px;
                     font-weight:700;color:#e0f0ff;letter-spacing:2px;">
            📡 nPOI DAS CONFIGURATOR
        </span><br>
        <span style="font-size:12px;color:#5a8aaa;letter-spacing:3px;">
            DISTRIBUTED ANTENNA SYSTEM · RACK 19" 1U · 8 PORTS RF PAR nPOI
        </span>
    </div>
    """, unsafe_allow_html=True)

    # ── Sidebar : paramètres ──────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⚙️ Paramètres")

        nb_secteurs = st.slider("Nombre de secteurs", 1, 12, 3)

        st.markdown("**Opérateurs**")
        operateurs_defaut = ["OPE1", "OPE2"]
        operateurs = st.multiselect(
            "Sélectionner / saisir les opérateurs",
            options=["OPE1", "OPE2", "OPE3", "Orange", "SFR", "Bouygues",
                     "Free", "VOO", "Proximus", "BASE"],
            default=operateurs_defaut,
        )
        op_custom = st.text_input("Ajouter un opérateur personnalisé", placeholder="Ex: Telenet")
        if op_custom and op_custom not in operateurs:
            operateurs.append(op_custom)

        st.markdown("**Fréquences (MHz)**")
        frequences = st.multiselect(
            "Sélectionner les bandes",
            options=FREQUENCES_ORDRE,
            default=["700/800", "1800", "2100"],
        )

        mimo = st.checkbox("MIMO 2×2 (double chaîne V/H)", value=False)

        tri = st.radio(
            "Trier les ports par",
            options=["Par fréquence", "Par opérateur"],
            index=0,
        )

        st.markdown("---")
        generer = st.button("🔧 Générer la configuration")

    # ── Corps principal ───────────────────────────────────────────────────────
    if not operateurs:
        st.warning("⚠️ Veuillez sélectionner au moins un opérateur.")
        return
    if not frequences:
        st.warning("⚠️ Veuillez sélectionner au moins une fréquence.")
        return

    # Calcul en temps réel
    ports  = construire_ports(nb_secteurs, operateurs, frequences, mimo, tri)
    npois  = grouper_en_npoi(ports)

    # ── Métriques ─────────────────────────────────────────────────────────────
    col1, col2, col3, col4 = st.columns(4)
    ports_total = len(ports)
    ports_libres = len(npois) * 8 - ports_total

    col1.metric("nPOI nécessaires", len(npois))
    col2.metric("Ports utilisés", ports_total)
    col3.metric("Ports libres", ports_libres)
    col4.metric("Unités de rack", f"{len(npois)} U")

    st.markdown("---")

    # ── Visualisation ─────────────────────────────────────────────────────────
    st.markdown("### 🗂️ Plan d'organisation")

    # Définition des couleurs HTML pour l'affichage
    freq_html = {
        "700/800": ("#1a3a5c", "#7ec8e3"),
        "900":     ("#1a4a2a", "#7ed9a0"),
        "1800":    ("#4a1a1a", "#e37e7e"),
        "2100":    ("#3a1a4a", "#c07ee3"),
        "2600":    ("#4a3a1a", "#e3c07e"),
        "3500":    ("#1a3a4a", "#7ec0e3"),
    }
    op_html = [
        ("#1a3a5c", "#7ec8e3"),
        ("#1a4a2a", "#7ed9a0"),
        ("#4a1a1a", "#e37e7e"),
        ("#3a1a4a", "#c07ee3"),
        ("#4a3a1a", "#e3c07e"),
        ("#1a3a4a", "#7ec0e3"),
    ]

    def get_color(port):
        if port is None:
            return "#2a2a2a", "#888888"
        if tri == "Par fréquence":
            return freq_html.get(port["frequence"], ("#2a2a2a", "#888888"))
        else:
            idx = operateurs.index(port["operateur"]) if port["operateur"] in operateurs else 0
            return op_html[idx % len(op_html)]

    # Affichage sur 2 colonnes
    cols = st.columns(2)
    for i, npoi in enumerate(npois):
        col = cols[i % 2]
        with col:
            col.markdown(
                f'<div class="npoi-header">🔌 nPOI {i+1} &nbsp;·&nbsp; '
                f'{len(npoi)}/8 ports</div>',
                unsafe_allow_html=True,
            )
            for p in range(8):
                port = npoi[p] if p < len(npoi) else None
                bg, fg = get_color(port)
                label = label_port(port)
                col.markdown(
                    f'<div class="port-card" style="'
                    f'background:{bg};color:{fg};border-color:{fg}40;">'
                    f'<b>P{p+1}</b> &nbsp; {label}'
                    f'</div>',
                    unsafe_allow_html=True,
                )

    # ── Légende ───────────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 🎨 Légende")
    legend_cols = st.columns(len(frequences) if tri == "Par fréquence" else len(operateurs))

    if tri == "Par fréquence":
        for i, freq in enumerate(frequences):
            bg, fg = freq_html.get(freq, ("#2a2a2a", "#888888"))
            legend_cols[i].markdown(
                f'<div style="background:{bg};color:{fg};padding:8px 12px;'
                f'border-radius:4px;text-align:center;font-size:13px;">'
                f'<b>{freq} MHz</b></div>',
                unsafe_allow_html=True,
            )
    else:
        for i, op in enumerate(operateurs):
            bg, fg = op_html[i % len(op_html)]
            legend_cols[i].markdown(
                f'<div style="background:{bg};color:{fg};padding:8px 12px;'
                f'border-radius:4px;text-align:center;font-size:13px;">'
                f'<b>{op}</b></div>',
                unsafe_allow_html=True,
            )

    # ── Export Excel ──────────────────────────────────────────────────────────
    st.markdown("---")
    if generer or True:  # Toujours proposer le téléchargement
        excel_buf = generer_excel(
            nb_secteurs, operateurs, frequences, mimo, tri, npois
        )
        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=excel_buf,
            file_name="nPOI_DAS_Configuration.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
