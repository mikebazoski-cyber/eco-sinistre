import pandas as pd
import unicodedata
import streamlit as st
from io import BytesIO

# =========================================================
# CONFIGURATION DE LA PAGE
# =========================================================
st.set_page_config(
    page_title="CarbonRepair Advisor",
    page_icon="🌿",
    layout="wide"
)

# =========================================================
# CHEMINS DES FICHIERS
# =========================================================
CARBON_FILE = "carbon_data.html"
COMPANIES_FILE = "entreprises.xlsx"

# =========================================================
# PARAMÈTRES MÉTIER
# =========================================================
REPL_MAP = {
    r"\'ea": "ê",
    r"\'e9": "é",
    r"\'e8": "è",
    r"\'b": "",
    r"\'ef": "ï",
    r"\'e7": "ç",
    r"\'e2": "â",
    r"\'9c": "œ",
    r"\'e0": "à",
    r"\'ee": "î",
}

SELECTOR_MAP = {
    "Menuiseries extérieures": "Extérieure",
    "Menuiserie intérieure": "Intérieure",
    "Revêtements de sol": "Sol",
    "Revêtements murs et plafonds": "Murs et plafonds",
    "Charpente - Ossature": "Charpente - Ossature",
    "Maçonnerie - Gros œuvre": "Maçonnerie - Gros œuvre",
    "Plomberie": "Plomberie",
    "Electricité": "Electricité",
    "Chauffage - Ventilation - Climatisation": "Chauffage - Ventilation - Climatisation",
}

CATEGORY_MERGE_MAP = {
    "Revêtements de sol": "Revêtements intérieurs",
    "Revêtements murs et plafonds": "Revêtements intérieurs",
    "Menuiseries extérieures": "Menuiseries",
    "Menuiserie intérieure": "Menuiseries",
    "Charpente - Ossature": "Structure",
    "Maçonnerie - Gros œuvre": "Structure",
    "Plomberie": "Réseaux techniques",
    "Electricité": "Réseaux techniques",
    "Chauffage - Ventilation - Climatisation": "Réseaux techniques",
}

LOW_CARBON_KEYWORDS = [
    "bas carbone",
    "chaume",
    "végétalisée",
    "biosourcé",
    "biosourcée",
    "laine",
    "chanvre",
    "ouate de cellulose",
]

# =========================================================
# SESSION STATE
# =========================================================
if "basket" not in st.session_state:
    st.session_state.basket = []

# =========================================================
# FONCTIONS UTILITAIRES
# =========================================================
def normalize_text(value: str) -> str:
    if pd.isna(value):
        return ""
    text = str(value).lower().strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return text

NORMALIZED_LOW_CARBON_KEYWORDS = [normalize_text(x) for x in LOW_CARBON_KEYWORDS]

def is_low_carbon_option(row: pd.Series) -> bool:
    text = f"{row.get('Sous_categorie', '')} {row.get('Produit_process', '')}"
    text_norm = normalize_text(text)
    keyword_match = any(keyword in text_norm for keyword in NORMALIZED_LOW_CARBON_KEYWORDS)

    emissions = row.get("Emissions_CO2")
    emissions_rule = pd.notna(emissions) and float(emissions) <= 0

    return keyword_match or emissions_rule

def split_categories(value):
    if pd.isna(value):
        return []
    text = str(value).replace("|", ";").replace(",", ";")
    return [x.strip() for x in text.split(";") if x.strip()]

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Chiffrage")
    return output.getvalue()

# =========================================================
# CHARGEMENT DES DONNÉES
# =========================================================
@st.cache_data
def load_carbon_df(html_path: str) -> pd.DataFrame:
    tables = pd.read_html(html_path)
    if not tables:
        raise ValueError("Aucun tableau trouvé dans carbon_data.html")

    df = tables[0].copy()

    df.columns = [
        "Categorie",
        "Sous_categorie",
        "Produit_process",
        "Unite",
        "Type_prestation",
        "Prestation",
        "Emissions_CO2",
    ]

    df = df.iloc[1:].reset_index(drop=True)

    for col in df.columns:
        if df[col].dtype == object:
            s = df[col].astype(str)
            for pat, repl in REPL_MAP.items():
                s = s.str.replace(pat, repl, regex=False)
            df[col] = s

    df["Emissions_CO2"] = pd.to_numeric(df["Emissions_CO2"], errors="coerce")
    df["Categorie_old"] = df["Categorie"]
    df["Selector"] = df["Categorie"].map(SELECTOR_MAP)
    df["Categorie"] = df["Categorie"].replace(CATEGORY_MERGE_MAP)

    return df

@st.cache_data
def load_companies_df(file_path: str) -> pd.DataFrame:
    if file_path.endswith(".xlsx") or file_path.endswith(".xls"):
        df = pd.read_excel(file_path)
    elif file_path.endswith(".csv"):
        df = pd.read_csv(file_path, encoding="utf-8")
    else:
        raise ValueError("Format du fichier entreprises non supporté.")

    df.columns = df.columns.astype(str).str.strip()

    rename_map = {
        "Entreprise": "Entreprise",
        "Description": "Description",
        "Spécificités": "Specificites",
        "Services": "Specificites",
        "Type de solution": "Type_solution",
        "Type de solution : technique, pratique ou les deux ?": "Type_solution",
        "indication du volume d'émissions qu'ils permettent d'économiser": "Volume_emissions_evitees",
        "Catégorie d’outil associée": "Categorie_outil",
        "Catégorie d'outil associée": "Categorie_outil",
        "Catégorie": "Categorie_outil",
        "Pays de couverture": "Pays_couverture",
        "Région de couverture": "Pays_couverture",
        "Siège(s) social(s)": "Siege",
        "Région du/des siège(s) social(s)": "Region_siege",
        "Lien": "Lien",
        "Link": "Lien",
        "Commentaires": "Commentaires",
        "Category List": "Category_List",
        "Subcategory List": "Subcategory_List",
    }

    df = df.rename(columns=rename_map)

    required_basic = ["Entreprise", "Categorie_outil"]
    for col in required_basic:
        if col not in df.columns:
            raise ValueError(f"Colonne obligatoire absente dans le fichier entreprises : {col}")

    df["Categorie_outil_liste"] = df["Categorie_outil"].apply(split_categories)

    return df

def build_candidates(filtered_df: pd.DataFrame) -> pd.DataFrame:
    candidates = (
        filtered_df[
            [
                "Categorie",
                "Categorie_old",
                "Selector",
                "Sous_categorie",
                "Produit_process",
                "Unite",
                "Type_prestation",
                "Prestation",
                "Emissions_CO2",
            ]
        ]
        .dropna(subset=["Produit_process", "Emissions_CO2"])
        .drop_duplicates()
        .copy()
    )

    if candidates.empty:
        return candidates

    candidates["Option_famille"] = candidates.apply(
        lambda row: "Option bas carbone" if is_low_carbon_option(row) else "Standard",
        axis=1,
    )

    candidates = candidates.sort_values(
        ["Option_famille", "Emissions_CO2", "Produit_process"],
        ascending=[True, True, True],
    ).reset_index(drop=True)

    return candidates

def make_option_table(option_df: pd.DataFrame) -> pd.DataFrame:
    if option_df.empty:
        return option_df

    table = option_df.copy()
    table["Émissions spécifiques (kg CO₂ / unité)"] = table["Emissions_CO2"].astype(float).round(2)

    table = table[
        ["Produit_process", "Unite", "Émissions spécifiques (kg CO₂ / unité)"]
    ].rename(
        columns={
            "Produit_process": "Produit / process",
            "Unite": "Unité",
        }
    )

    return table.reset_index(drop=True)

def filter_companies_by_category(companies_df: pd.DataFrame, selected_category: str) -> pd.DataFrame:
    result = companies_df[
        companies_df["Categorie_outil_liste"].apply(lambda cats: selected_category in cats)
    ].copy()

    cols = [
        "Entreprise",
        "Description",
        "Specificites",
        "Type_solution",
        "Pays_couverture",
        "Siege",
        "Region_siege",
        "Lien",
    ]
    cols = [c for c in cols if c in result.columns]

    return result[cols].reset_index(drop=True)

# =========================================================
# CHARGEMENT
# =========================================================
carbon_df = load_carbon_df(CARBON_FILE)
companies_df = load_companies_df(COMPANIES_FILE)

# =========================================================
# STYLE
# =========================================================
st.markdown("""
<style>
.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}
.card {
    background: #ffffff;
    color: #1f1f1f;
    border: 1px solid #d9e2f3;
    border-radius: 12px;
    padding: 16px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08);
}
.header-card {
    background: linear-gradient(90deg, #1f4e79, #2f75b5);
    color: white;
    padding: 20px;
    border-radius: 12px;
    margin-bottom: 15px;
}
.info-card {
    background: #f9fcff;
    color: #1f1f1f;
    border: 1px solid #d9e2f3;
    border-radius: 12px;
    padding: 16px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08);
}
.section-title {
    color: #1f4e79;
    margin-top: 1rem;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# EN-TÊTE
# =========================================================
st.markdown("""
<div class="header-card">
    <h2 style="margin:0;">CarbonRepair Advisor</h2>
    <p style="margin:10px 0 0 0; line-height:1.5;">
        Bienvenue dans cet outil d’aide au chiffrage sinistre bas carbone.
        Il permet de sélectionner une solution de réparation, de comparer une option standard
        avec une alternative bas carbone, d’estimer les émissions de CO₂ associées
        et d’identifier les entreprises liées à la catégorie sélectionnée.
    </p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="info-card">
    <b>Parcours utilisateur</b><br>
    L’outil est structuré en quatre étapes : sélection du poste sinistré, estimation des émissions et comparaison des solutions,
    identification des entreprises liées à la catégorie choisie, puis récapitulatif du chiffrage.
</div>
""", unsafe_allow_html=True)

# =========================================================
# ÉTAPE 1
# =========================================================
st.markdown("<h3 class='section-title'>Étape 1 — Sélection du poste sinistré</h3>", unsafe_allow_html=True)
st.markdown("""
Dans cette première étape, vous définissez le poste sinistré en sélectionnant la catégorie,
le niveau de détail technique et la prestation concernée.
Cette sélection permet d’orienter l’outil vers les solutions les plus pertinentes.
""")

categories = sorted(carbon_df["Categorie"].dropna().unique().tolist())
selected_category = st.selectbox("Catégorie", categories)

d1 = carbon_df[carbon_df["Categorie"] == selected_category].copy()

selector_options = sorted([x for x in d1["Selector"].dropna().unique().tolist() if x != ""])
if selector_options:
    selected_selector = st.selectbox("Sélecteur", selector_options)
    d2 = d1[d1["Selector"] == selected_selector].copy()
else:
    selected_selector = ""
    d2 = d1.copy()

sous_cat_options = sorted(d2["Sous_categorie"].dropna().unique().tolist())
selected_sous_cat = st.selectbox("Sous-catégorie", sous_cat_options)

d3 = d2[d2["Sous_categorie"] == selected_sous_cat].copy()

type_prest_options = sorted(d3["Type_prestation"].dropna().unique().tolist())
selected_type_prest = st.selectbox("Type de prestation", type_prest_options)

d4 = d3[d3["Type_prestation"] == selected_type_prest].copy()

prest_options = sorted(d4["Prestation"].dropna().unique().tolist())
selected_prest = st.selectbox("Prestation", prest_options)

d5 = d4[d4["Prestation"] == selected_prest].copy()

current_candidates = build_candidates(d5)

standard_df = current_candidates[current_candidates["Option_famille"] == "Standard"].reset_index(drop=True)
low_carbon_df = current_candidates[current_candidates["Option_famille"] == "Option bas carbone"].reset_index(drop=True)

available_families = []
if not standard_df.empty:
    available_families.append("Standard")
if not low_carbon_df.empty:
    available_families.append("Option bas carbone")

current_selected_row = None
qty = 0.0
selected_family = None

if available_families:
    selected_family = st.radio("Type d’option", available_families, horizontal=True)

    active_df = current_candidates[
        current_candidates["Option_famille"] == selected_family
    ].reset_index(drop=True)

    active_df = active_df.copy()
    active_df["label"] = active_df.apply(
        lambda row: f"{row['Produit_process']} — {float(row['Emissions_CO2']):.2f} kg CO₂ / {row['Unite']}",
        axis=1
    )

    selected_label = st.selectbox("Produit / process", active_df["label"].tolist())
    current_selected_row = active_df[active_df["label"] == selected_label].iloc[0]

    qty = st.number_input("Quantité", min_value=0.0, value=1.0, step=1.0)
else:
    st.warning("Aucune option disponible pour les filtres sélectionnés.")

# =========================================================
# ÉTAPE 2
# =========================================================
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("<h3 class='section-title'>Étape 2 — Estimation des émissions et comparaison des solutions</h3>", unsafe_allow_html=True)
st.markdown("""
Une fois la solution choisie, l’outil estime les émissions de CO₂ associées à la quantité renseignée.
Il affiche d’abord les solutions standard, puis les alternatives bas carbone.
Un comparateur d’impact met également en évidence le gain carbone potentiel
et formule une recommandation.
""")

emissions_per_unit = None
emissions_total = None
unit = ""

if current_selected_row is not None:
    unit = str(current_selected_row["Unite"]) if pd.notna(current_selected_row["Unite"]) else ""
    emissions_per_unit = float(current_selected_row["Emissions_CO2"])
    emissions_total = emissions_per_unit * float(qty)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div class="card">
            <div style="font-size:15px; font-weight:600; color:#1f4e79; margin-bottom:8px;">
                Émissions spécifiques
            </div>
            <div style="font-size:18px;">
                {emissions_per_unit:.2f} kg CO₂ / {unit}
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class="card">
            <div style="font-size:15px; font-weight:600; color:#1f4e79; margin-bottom:8px;">
                Émissions totales
            </div>
            <div style="font-size:18px;">
                {emissions_total:.2f} kg CO₂
            </div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("#### Comparateur d’impact")
st.markdown("""
Lorsque des solutions standard et bas carbone existent pour un même poste,
l’outil compare automatiquement leurs impacts afin de mettre en évidence
le gain carbone potentiel et d’aider à la prise de décision.
""")

if not standard_df.empty and not low_carbon_df.empty:
    best_standard = standard_df.sort_values("Emissions_CO2", ascending=True).iloc[0]
    best_low_carbon = low_carbon_df.sort_values("Emissions_CO2", ascending=True).iloc[0]

    standard_total = float(best_standard["Emissions_CO2"]) * float(qty)
    low_carbon_total = float(best_low_carbon["Emissions_CO2"]) * float(qty)
    gain_absolute = standard_total - low_carbon_total
    reduction_pct = (gain_absolute / standard_total * 100) if standard_total > 0 else 0.0

    if gain_absolute > 0:
        recommendation = "Privilégier l’alternative bas carbone."
        recommendation_color = "#2e7d32"
    elif gain_absolute < 0:
        recommendation = "La solution standard présente ici un impact carbone inférieur."
        recommendation_color = "#c62828"
    else:
        recommendation = "Les deux solutions présentent un impact équivalent selon les données disponibles."
        recommendation_color = "#7f6000"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Émissions standard", f"{standard_total:.2f} kg CO₂")
    c2.metric("Émissions bas carbone", f"{low_carbon_total:.2f} kg CO₂")
    c3.metric("Gain carbone absolu", f"{gain_absolute:.2f} kg CO₂")
    c4.metric("Réduction", f"{reduction_pct:.1f} %")

    st.markdown(f"""
    <div class="card" style="border-left:5px solid {recommendation_color};">
        <b>Recommandation :</b> <span style="color:{recommendation_color};">{recommendation}</span><br><br>
        <b>Référence standard :</b> {best_standard['Produit_process']}<br>
        <b>Alternative bas carbone :</b> {best_low_carbon['Produit_process']}
    </div>
    """, unsafe_allow_html=True)
else:
    st.info("Comparaison d’impact indisponible : au moins une solution standard et une solution bas carbone sont nécessaires.")

# Boutons
b1, b2, b3 = st.columns(3)

with b1:
    if st.button("Ajouter au chiffrage", use_container_width=True):
        if current_selected_row is not None and emissions_per_unit is not None and emissions_total is not None:
            st.session_state.basket.append(
                {
                    "Categorie": str(current_selected_row["Categorie"]),
                    "Selector": "" if not selector_options else str(selected_selector),
                    "Sous_categorie": str(current_selected_row["Sous_categorie"]),
                    "Type_prestation": str(current_selected_row["Type_prestation"]),
                    "Prestation": str(current_selected_row["Prestation"]),
                    "Option_famille": str(selected_family),
                    "Produit_process": str(current_selected_row["Produit_process"]),
                    "Unite": unit,
                    "Quantite": float(qty),
                    "Emissions_specifiques": float(emissions_per_unit),
                    "kg_CO2_total": float(emissions_total),
                }
            )
            st.success("Ligne ajoutée au chiffrage.")

with b2:
    if st.button("Retirer la dernière ligne", use_container_width=True):
        if st.session_state.basket:
            st.session_state.basket.pop()
            st.success("Dernière ligne retirée.")

with b3:
    if st.button("Vider le chiffrage", use_container_width=True):
        st.session_state.basket = []
        st.success("Chiffrage vidé.")

st.markdown("#### Comparaison des solutions disponibles")

st.markdown("##### Solutions standard")
if standard_df.empty:
    st.write("Aucune solution standard trouvée.")
else:
    st.dataframe(make_option_table(standard_df), use_container_width=True)

st.markdown("##### Solutions bas carbone")
if low_carbon_df.empty:
    st.write("Aucune solution bas carbone trouvée.")
else:
    st.dataframe(make_option_table(low_carbon_df), use_container_width=True)

# =========================================================
# ÉTAPE 3
# =========================================================
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("<h3 class='section-title'>Étape 3 — Entreprises liées à la catégorie sélectionnée</h3>", unsafe_allow_html=True)
st.markdown("""
Cette étape présente les entreprises liées à la catégorie sélectionnée.
Elle permet d’identifier rapidement des acteurs pertinents, leurs spécificités,
leur zone de couverture, leur siège social et leurs informations principales.
""")

company_result = filter_companies_by_category(companies_df, selected_category)

st.markdown(f"""
<div class="card">
    <div style="font-size:15px; font-weight:600; color:#1f4e79; margin-bottom:8px;">
        Informations de correspondance
    </div>
    <div><b>Catégorie sélectionnée :</b> {selected_category}</div>
    <div><b>Entreprises identifiées :</b> {len(company_result)}</div>
</div>
""", unsafe_allow_html=True)

if company_result.empty:
    st.write("Aucune entreprise trouvée pour cette catégorie.")
else:
    st.dataframe(company_result, use_container_width=True)

# =========================================================
# ÉTAPE 4
# =========================================================
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown("<h3 class='section-title'>Étape 4 — Récapitulatif du chiffrage</h3>", unsafe_allow_html=True)
st.markdown("""
Le récapitulatif du chiffrage rassemble toutes les lignes ajoutées,
les quantités sélectionnées, les émissions associées à chaque ligne
et le total global estimé en CO₂.
""")

if not st.session_state.basket:
    st.write("Aucune ligne ajoutée pour le moment.")
else:
    basket_df = pd.DataFrame(st.session_state.basket)

    basket_display = basket_df.copy()
    basket_display["Quantite"] = basket_display["Quantite"].round(2)
    basket_display["Emissions_specifiques"] = basket_display["Emissions_specifiques"].round(2)
    basket_display["kg_CO2_total"] = basket_display["kg_CO2_total"].round(2)

    basket_display = basket_display.rename(
        columns={
            "Selector": "Sélecteur",
            "Sous_categorie": "Sous-catégorie",
            "Type_prestation": "Type de prestation",
            "Option_famille": "Type d’option",
            "Produit_process": "Produit / process",
            "Unite": "Unité",
            "Quantite": "Quantité",
            "Emissions_specifiques": "Émissions spécifiques (kg CO₂ / unité)",
            "kg_CO2_total": "kg CO₂ total",
        }
    )

    display_cols = [
        "Categorie",
        "Sélecteur",
        "Sous-catégorie",
        "Type de prestation",
        "Prestation",
        "Type d’option",
        "Produit / process",
        "Unité",
        "Quantité",
        "Émissions spécifiques (kg CO₂ / unité)",
        "kg CO₂ total",
    ]

    st.dataframe(basket_display[display_cols], use_container_width=True)

    total = float(basket_df["kg_CO2_total"].sum())
    st.markdown(f"**Total global estimé :** {total:.2f} kg CO₂")

    csv_data = basket_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="Télécharger le chiffrage en CSV",
        data=csv_data,
        file_name="chiffrage_sinistre.csv",
        mime="text/csv",
    )

    excel_bytes = to_excel_bytes(basket_df)
    st.download_button(
        label="Télécharger le chiffrage en Excel",
        data=excel_bytes,
        file_name="chiffrage_sinistre.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )