import io
import random
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Outil Santé - Génération des Quantités",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================
# STYLE
# =========================
st.markdown("""
<style>
.main {
    background: #f7f9fc;
}
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
}
.hero-card {
    background: linear-gradient(135deg, #16324f, #245b8f);
    padding: 22px 24px;
    border-radius: 18px;
    color: white;
    box-shadow: 0 10px 25px rgba(0,0,0,0.10);
    margin-bottom: 18px;
}
.small-card {
    background: white;
    padding: 16px;
    border-radius: 16px;
    box-shadow: 0 8px 20px rgba(0,0,0,0.06);
    border: 1px solid #edf1f7;
}
.section-title {
    font-size: 1.15rem;
    font-weight: 700;
    margin-bottom: 8px;
    color: #1c3557;
}
.kpi-box {
    background: white;
    padding: 18px;
    border-radius: 16px;
    border: 1px solid #e8eef5;
    box-shadow: 0 8px 20px rgba(0,0,0,0.05);
    text-align: center;
}
.kpi-label {
    color: #6b7788;
    font-size: 0.95rem;
}
.kpi-value {
    color: #12304d;
    font-size: 1.45rem;
    font-weight: 800;
}
.stTabs [data-baseweb="tab-list"] {
    gap: 10px;
}
.stTabs [data-baseweb="tab"] {
    background: #eef4fb;
    border-radius: 12px;
    padding: 10px 16px;
}
.stTabs [aria-selected="true"] {
    background: #1f5f99 !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

# =========================
# MOT DE PASSE
# =========================
APP_PASSWORD = "EDDAQAQ2026"

def login():
    st.markdown('<div class="hero-card"><h2>🔐 Accès sécurisé</h2><p>Entrez le mot de passe pour accéder à l’outil.</p></div>', unsafe_allow_html=True)
    pwd = st.text_input("Mot de passe", type="password")
    if st.button("Se connecter", use_container_width=True):
        if pwd == APP_PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Mot de passe incorrect.")

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    login()
    st.stop()

# =========================
# DONNÉES PAR RUBRIQUE
# =========================
MONTHS = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
]

YEARS = ["Année 1", "Année 2", "Année 3", "Année 4", "Année 5"]

SECTIONS = {
    "Centre de radiologie": [
        ("RX PHOTO", 50),
        ("RX(face)", 200),
        ("RX face et profil", 300),
        ("Echo Abdo", 500),
        ("Echo parties molles", 600),
        ("Echo cervical", 600),
        ("Echo mammaire", 600),
        ("Mammo", 600),
        ("Mammo+Echo", 800),
        ("Scanner cérébral", 1500),
        ("Scanner abdominal", 1500),
        ("TDM FACE", 1500),
        ("TDM SINUS", 1500),
        ("TDM cervical/dorsal/lombaire", 1500),
        ("TDM bassin", 1500),
        ("TDM hanche/genou/épaule/coude", 1500),
        ("Scanner thoracique", 1500),
        ("TDM rocher", 2000),
        ("Scanner abdomino pelvien", 2000),
        ("Scanner thoraco abdomino pelvien", 3000),
        ("Angio scanner", 3000),
    ],
    "Laboratoire": [
        ("NFS", 80),
        ("CRP", 90),
        ("Glycémie", 40),
        ("HbA1c", 120),
        ("Bilan hépatique", 180),
        ("Bilan rénal", 150),
        ("Ionogramme", 130),
        ("TSH", 140),
        ("Ferritine", 160),
        ("Vitamine D", 220),
        ("ECBU", 70),
        ("Coproculture", 110),
        ("Bêta HCG", 100),
        ("Bilan lipidique", 150),
        ("Sérologie", 190),
    ],
    "Clinique": [
        ("Consultation générale", 250),
        ("Consultation spécialisée", 400),
        ("Échographie clinique", 350),
        ("Acte infirmier", 120),
        ("Petite chirurgie", 900),
        ("Hospitalisation jour", 700),
        ("Hospitalisation nuit", 1200),
        ("Monitoring", 180),
        ("Pansement spécialisé", 140),
        ("Injection / perfusion", 100),
        ("Bilan pré-op", 300),
        ("Suivi postopératoire", 280),
    ]
}

DEFAULT_CA = {
    "Année 1": 4000000.0,
    "Année 2": 5000000.0,
    "Année 3": 6000000.0,
    "Année 4": 7200000.0,
    "Année 5": 7920000.0,
}

# =========================
# FONCTIONS
# =========================
def init_section_state(section_name):
    key_items = f"items_{section_name}"
    key_ca = f"ca_{section_name}"

    if key_items not in st.session_state:
        st.session_state[key_items] = pd.DataFrame(SECTIONS[section_name], columns=["Acte / Examen", "Prix Unitaire"])

    if key_ca not in st.session_state:
        st.session_state[key_ca] = DEFAULT_CA.copy()

def generate_quantities_for_year(ca_target, items_df, whole_numbers=False, seed=42):
    random.seed(seed)
    np.random.seed(seed)

    items = items_df["Acte / Examen"].tolist()
    prices = items_df["Prix Unitaire"].tolist()

    cells = []
    for item, price in zip(items, prices):
        for month in MONTHS:
            cells.append((item, month, float(price)))

    values = {}
    generated_ca = 0.0

    for i, (item, month, price) in enumerate(cells[:-1]):
        remaining_cells = len(cells) - i - 1
        remaining_ca = ca_target - generated_ca

        if remaining_ca <= 0:
            qty = 0.0
        else:
            avg_ca = remaining_ca / max(1, remaining_cells)
            factor = random.uniform(0.35, 1.85)
            proposed_ca = avg_ca * factor
            qty = proposed_ca / price if price > 0 else 0.0

            if whole_numbers:
                qty = max(0, int(round(qty)))
                while qty * price > remaining_ca and qty > 0:
                    qty -= 1

        values[(item, month)] = qty
        generated_ca += qty * price

    last_item, last_month, last_price = cells[-1]
    missing_ca = ca_target - generated_ca

    if last_price > 0:
        final_qty = missing_ca / last_price
        if whole_numbers and abs(final_qty - round(final_qty)) < 1e-9 and final_qty >= 0:
            final_qty = int(round(final_qty))
        elif final_qty < 0:
            final_qty = 0.0
    else:
        final_qty = 0.0

    values[(last_item, last_month)] = final_qty

    rows = []
    for item, price in zip(items, prices):
        row = {
            "Acte / Examen": item,
            "Prix Unitaire": price
        }
        total_qty = 0.0
        total_ca = 0.0

        for month in MONTHS:
            qty = values.get((item, month), 0.0)
            ca = qty * price
            row[f"Qté {month}"] = qty
            row[f"CA {month}"] = ca
            total_qty += qty
            total_ca += ca

        row["Qté Totale"] = total_qty
        row["CA Total"] = total_ca
        rows.append(row)

    detail_df = pd.DataFrame(rows)
    total_generated = detail_df["CA Total"].sum()
    return detail_df, total_generated

def build_monthly_summary(detail_df):
    rows = []
    for month in MONTHS:
        rows.append({
            "Mois": month,
            "Quantité Totale": detail_df[f"Qté {month}"].sum(),
            "CA Mensuel": detail_df[f"CA {month}"].sum()
        })
    return pd.DataFrame(rows)

def export_section_excel(section_name, ca_dict, all_results, all_monthly):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame({
            "Année": list(ca_dict.keys()),
            "CA cible": list(ca_dict.values())
        }).to_excel(writer, sheet_name="CA_Cible", index=False)

        for year_name, df in all_results.items():
            sheet_detail = (year_name.replace(" ", "_") + "_detail")[:31]
            sheet_month = (year_name.replace(" ", "_") + "_mensuel")[:31]
            df.to_excel(writer, sheet_name=sheet_detail, index=False)
            all_monthly[year_name].to_excel(writer, sheet_name=sheet_month, index=False)

    output.seek(0)
    return output

# =========================
# HEADER
# =========================
st.markdown("""
<div class="hero-card">
    <h1 style="margin-bottom:6px;">🏥 Outil Santé - Génération des Quantités</h1>
    <p style="margin:0;font-size:1.02rem;">
        Clinique • Laboratoire • Centre de Radiologie<br>
        Saisis les CA des 5 années, les actes et les prix unitaires, puis génère automatiquement les quantités mensuelles.
    </p>
</div>
""", unsafe_allow_html=True)

top1, top2, top3 = st.columns(3)
with top1:
    st.markdown('<div class="kpi-box"><div class="kpi-label">Rubriques</div><div class="kpi-value">3</div></div>', unsafe_allow_html=True)
with top2:
    st.markdown('<div class="kpi-box"><div class="kpi-label">Années gérées</div><div class="kpi-value">5</div></div>', unsafe_allow_html=True)
with top3:
    st.markdown('<div class="kpi-box"><div class="kpi-label">Mode</div><div class="kpi-value">Sécurisé</div></div>', unsafe_allow_html=True)

st.write("")

# =========================
# SIDEBAR
# =========================
st.sidebar.title("Navigation")
selected_section = st.sidebar.radio(
    "Choisir une rubrique",
    ["Clinique", "Laboratoire", "Centre de radiologie"]
)

whole_numbers = st.sidebar.checkbox("Quantités plutôt entières", value=True)
seed_value = st.sidebar.number_input("Seed aléatoire", min_value=0, value=42, step=1)

if st.sidebar.button("Se déconnecter", use_container_width=True):
    st.session_state["authenticated"] = False
    st.rerun()

init_section_state(selected_section)

items_key = f"items_{selected_section}"
ca_key = f"ca_{selected_section}"

# =========================
# LAYOUT PRINCIPAL
# =========================
tab1, tab2, tab3, tab4 = st.tabs([
    "1. Paramétrage",
    "2. Génération",
    "3. Résultats",
    "4. Export"
])

with tab1:
    left, right = st.columns([1, 1.6])

    with left:
        st.markdown('<div class="small-card">', unsafe_allow_html=True)
        st.markdown(f'<div class="section-title">CA cible - {selected_section}</div>', unsafe_allow_html=True)

        current_ca = st.session_state[ca_key]
        new_ca = {}

        for y in YEARS:
            new_ca[y] = st.number_input(
                y,
                min_value=0.0,
                value=float(current_ca[y]),
                step=100000.0,
                key=f"{selected_section}_{y}"
            )

        st.session_state[ca_key] = new_ca
        st.markdown('</div>', unsafe_allow_html=True)

    with right:
        st.markdown('<div class="small-card">', unsafe_allow_html=True)
        st.markdown(f'<div class="section-title">Actes / Examens - {selected_section}</div>', unsafe_allow_html=True)

        edited_df = st.data_editor(
            st.session_state[items_key],
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_{selected_section}"
        )

        edited_df["Acte / Examen"] = edited_df["Acte / Examen"].astype(str).str.strip()
        edited_df["Prix Unitaire"] = pd.to_numeric(edited_df["Prix Unitaire"], errors="coerce").fillna(0.0)
        edited_df = edited_df[edited_df["Acte / Examen"] != ""].reset_index(drop=True)

        st.session_state[items_key] = edited_df
        st.markdown('</div>', unsafe_allow_html=True)

with tab2:
    st.markdown('<div class="small-card">', unsafe_allow_html=True)
    st.markdown(f'<div class="section-title">Génération automatique - {selected_section}</div>', unsafe_allow_html=True)
    st.write("Clique sur le bouton pour générer les quantités mensuelles sur les 5 années.")

    if st.button("Générer les quantités", type="primary", use_container_width=True):
        items_df = st.session_state[items_key].copy()
        ca_dict = st.session_state[ca_key]

        if items_df.empty:
            st.error("Ajoute au moins une ligne d'acte ou d'examen.")
        elif (items_df["Prix Unitaire"] <= 0).any():
            st.error("Tous les prix unitaires doivent être supérieurs à 0.")
        else:
            all_results = {}
            all_monthly = {}
            controls = []

            for idx, year_name in enumerate(YEARS):
                detail_df, total_generated = generate_quantities_for_year(
                    ca_target=ca_dict[year_name],
                    items_df=items_df,
                    whole_numbers=whole_numbers,
                    seed=seed_value + idx
                )
                monthly_df = build_monthly_summary(detail_df)

                all_results[year_name] = detail_df
                all_monthly[year_name] = monthly_df

                gap = total_generated - ca_dict[year_name]
                controls.append({
                    "Année": year_name,
                    "CA cible": ca_dict[year_name],
                    "CA généré": total_generated,
                    "Écart": gap
                })

            st.session_state[f"results_{selected_section}"] = all_results
            st.session_state[f"monthly_{selected_section}"] = all_monthly
            st.session_state[f"controls_{selected_section}"] = pd.DataFrame(controls)

            st.success("Génération terminée avec succès.")
    st.markdown('</div>', unsafe_allow_html=True)

with tab3:
    results_key = f"results_{selected_section}"
    monthly_key = f"monthly_{selected_section}"
    controls_key = f"controls_{selected_section}"

    if results_key not in st.session_state:
        st.info("Aucun résultat encore généré. Va dans l’onglet Génération.")
    else:
        controls_df = st.session_state[controls_key]
        all_results = st.session_state[results_key]
        all_monthly = st.session_state[monthly_key]

        st.markdown('<div class="small-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Contrôle global</div>', unsafe_allow_html=True)
        st.dataframe(controls_df, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.write("")

        year_choice = st.segmented_control(
            "Choisir une année",
            options=YEARS,
            selection_mode="single",
            default="Année 1",
            key=f"segment_{selected_section}"
        )

        if year_choice is None:
            year_choice = "Année 1"

        detail_df = all_results[year_choice]
        monthly_df = all_monthly[year_choice]

        k1, k2, k3 = st.columns(3)
        with k1:
            st.markdown(f'<div class="kpi-box"><div class="kpi-label">Rubrique</div><div class="kpi-value">{selected_section}</div></div>', unsafe_allow_html=True)
        with k2:
            st.markdown(f'<div class="kpi-box"><div class="kpi-label">Année</div><div class="kpi-value">{year_choice}</div></div>', unsafe_allow_html=True)
        with k3:
            ca_total = monthly_df["CA Mensuel"].sum()
            st.markdown(f'<div class="kpi-box"><div class="kpi-label">CA recalculé</div><div class="kpi-value">{ca_total:,.0f}</div></div>', unsafe_allow_html=True)

        st.write("")

        c1, c2 = st.columns([1.1, 1])
        with c1:
            st.markdown('<div class="small-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">CA mensuel</div>', unsafe_allow_html=True)
            fig_month = px.bar(
                monthly_df,
                x="Mois",
                y="CA Mensuel",
                text_auto=".2s",
                title=""
            )
            fig_month.update_layout(height=420)
            st.plotly_chart(fig_month, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with c2:
            st.markdown('<div class="small-card">', unsafe_allow_html=True)
            st.markdown('<div class="section-title">Répartition par acte / examen</div>', unsafe_allow_html=True)
            exam_chart = detail_df[["Acte / Examen", "CA Total"]].sort_values("CA Total", ascending=False)
            fig_exam = px.bar(
                exam_chart,
                x="Acte / Examen",
                y="CA Total",
                title=""
            )
            fig_exam.update_layout(height=420)
            st.plotly_chart(fig_exam, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        st.write("")

        st.markdown('<div class="small-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Synthèse mensuelle</div>', unsafe_allow_html=True)
        st.dataframe(monthly_df, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.write("")

        st.markdown('<div class="small-card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Détail complet des quantités et du CA</div>', unsafe_allow_html=True)
        st.dataframe(detail_df, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

with tab4:
    results_key = f"results_{selected_section}"
    monthly_key = f"monthly_{selected_section}"

    st.markdown('<div class="small-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Export Excel</div>', unsafe_allow_html=True)

    if results_key not in st.session_state:
        st.info("Aucun résultat à exporter pour le moment.")
    else:
        excel_file = export_section_excel(
            selected_section,
            st.session_state[ca_key],
            st.session_state[results_key],
            st.session_state[monthly_key]
        )

        st.download_button(
            label=f"Télécharger l'export Excel - {selected_section}",
            data=excel_file,
            file_name=f"{selected_section.lower().replace(' ', '_')}_quantites.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    st.markdown('</div>', unsafe_allow_html=True)