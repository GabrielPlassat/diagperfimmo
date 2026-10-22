"""Application Streamlit : Diag Perf'Immo (PDF) -> modèle d'indicateurs (Excel), via l'API Albert."""
import hashlib
import os
from datetime import datetime

import pandas as pd
import streamlit as st

import perfimmo as pi

st.set_page_config(page_title="Extraction Diag Perf'Immo", page_icon="🏢", layout="wide")


# ---------------------------------------------------------------- utilitaires
def secret(nom, defaut=""):
    try:
        return st.secrets.get(nom, os.environ.get(nom, defaut))
    except Exception:
        return os.environ.get(nom, defaut)


@st.cache_data(show_spinner=False, max_entries=20)
def lire_pdf_cache(contenu: bytes) -> list[str]:
    return pi.lire_pdf(contenu)


@st.cache_data(show_spinner=False, ttl=3600)
def modeles_cache(cle: str, base_url: str) -> list[str]:
    return pi.lister_modeles(pi.client_albert(cle, base_url))


def empreinte(b: bytes) -> str:
    return hashlib.sha1(b).hexdigest()[:10]


ss = st.session_state
ss.setdefault("resultats", {})   # nom PDF -> {ligne: Resultat}
ss.setdefault("pdfs", {})        # nom PDF -> bytes

# ---------------------------------------------------------------- barre latérale
with st.sidebar:
    st.header("⚙️ Réglages")
    cle = secret("ALBERT_API_KEY")
    if cle:
        st.success("Clé API Albert chargée depuis les secrets")
    else:
        cle = st.text_input("Clé API Albert", type="password",
                            help="Créez une clé sur albert.sites.beta.gouv.fr. "
                                 "Pour ne plus la saisir : Streamlit Cloud > Settings > Secrets.")
    base_url = secret("ALBERT_BASE_URL", pi.ALBERT_BASE_URL)

    modeles = modeles_cache(cle, base_url) if cle else pi.MODELES_PREFERES
    modele = st.selectbox("Modèle", modeles, index=0,
                          help="gpt-oss-120b (alias openweight-large) donne en général les meilleurs résultats.")
    mode = st.radio(
        "Texte envoyé au modèle",
        ["complet", "cible"],
        format_func=lambda m: {"complet": "Rapport complet (plus fiable)",
                               "cible": "Pages ciblées (économise les quotas)"}[m],
    )
    paralleles = st.slider("Requêtes simultanées", 1, 4, 2,
                           help="Baissez à 1 si vous recevez des erreurs de quota (429).")

    st.divider()
    modele_xlsx = st.file_uploader("Modèle Excel (facultatif)", type=["xlsx"],
                                   help="Par défaut : « modèle indicateurs perfimmo - enrichis ». "
                                        "La colonne A doit contenir les indicateurs.")
    st.caption("Consignes par indicateur : fichier `consignes.csv` du dépôt GitHub.")

modele_bytes = modele_xlsx.getvalue() if modele_xlsx else pi.MODELE_PAR_DEFAUT.read_bytes()
indicateurs, feuille = pi.lire_indicateurs(modele_bytes)

# ---------------------------------------------------------------- en-tête
st.title("🏢 Extraction des rapports Diag Perf'Immo")
st.write(
    f"Déposez un ou plusieurs rapports PDF : l'IA **Albert** lit chaque rapport et remplit "
    f"les **{len(indicateurs)} indicateurs** du modèle Excel. Chaque valeur est colorée selon "
    f"le niveau de confiance et vous pouvez la corriger avant de télécharger le fichier."
)
with st.expander(f"Voir les indicateurs du modèle ({len(indicateurs)}, feuille « {feuille} »)"):
    st.dataframe(pd.DataFrame([{"Ligne": i.ligne, "Indicateur": i.libelle,
                                "Groupe": i.groupe, "Consigne": i.consigne} for i in indicateurs]),
                 hide_index=True, width="stretch")

# ---------------------------------------------------------------- 1. dépôt + extraction
st.subheader("1. Rapports à analyser")
fichiers = st.file_uploader("Rapports PDF", type=["pdf"], accept_multiple_files=True,
                            label_visibility="collapsed")

if fichiers:
    a_faire = [f for f in fichiers if f.name not in ss.resultats]
    col1, col2 = st.columns([1, 3])
    lancer = col1.button(f"▶️ Lancer l'extraction ({len(a_faire)} PDF)", type="primary",
                         disabled=not (cle and a_faire))
    if not cle:
        col2.warning("Saisissez d'abord votre clé API Albert dans la barre latérale.")
    elif not a_faire:
        col2.info("Tous les PDF déposés ont déjà été traités. Supprimez un résultat ci-dessous pour le relancer.")
    else:
        col2.caption("Compter environ 1 à 3 minutes par rapport.")

    if lancer:
        client = pi.client_albert(cle, base_url)
        for f in a_faire:
            contenu = f.getvalue()
            with st.status(f"📄 {f.name}", expanded=True) as statut:
                st.write("Lecture du PDF…")
                pages = lire_pdf_cache(contenu)
                st.write(f"{len(pages)} pages lues. Interrogation d'Albert ({modele})…")
                barre = st.progress(0.0)
                try:
                    res = pi.extraire_rapport(
                        client, modele, pages, indicateurs, mode=mode, paralleles=paralleles,
                        progression=lambda k, n: barre.progress(k / n, text=f"Groupe {k}/{n}"),
                    )
                except Exception as e:  # noqa: BLE001
                    statut.update(label=f"❌ {f.name}", state="error")
                    st.error(str(e))
                    continue
                ss.resultats[f.name] = res
                ss.pdfs[f.name] = contenu
                n_ok = sum(r.valeur is not None for r in res.values())
                statut.update(label=f"✅ {f.name} — {n_ok}/{len(res)} valeurs trouvées",
                              state="complete", expanded=False)
        st.rerun()

# ---------------------------------------------------------------- 2. relecture
if ss.resultats:
    st.subheader("2. Relecture")
    legende = " ".join(
        f'<span style="background:#{c};padding:2px 8px;border-radius:4px;color:#222">{k}</span>'
        for k, c in pi.COULEURS.items())
    st.markdown(legende, unsafe_allow_html=True)

    onglets = st.tabs(list(ss.resultats))
    for onglet, nom in zip(onglets, list(ss.resultats)):
        with onglet:
            res = ss.resultats[nom]
            comptes = pd.Series([r.confiance for r in res.values()]).value_counts()
            cols = st.columns(len(pi.CONFIANCES))
            for c, k in zip(cols, pi.CONFIANCES):
                c.metric(k.capitalize(), int(comptes.get(k, 0)))

            filtre = st.toggle("Afficher seulement les valeurs à vérifier (moyenne, faible)",
                               key=f"f_{nom}")
            df = pd.DataFrame([{
                "ligne": r.ligne, "Indicateur": r.indicateur,
                "Valeur": "" if r.valeur is None else str(r.valeur),
                "Confiance": r.confiance, "Page": r.page, "Extrait": r.extrait,
                "Commentaire / contrôles": " ; ".join(x for x in [r.commentaire] + r.controles if x),
            } for r in res.values()]).set_index("ligne")
            if filtre:
                df = df[df["Confiance"].isin(["moyenne", "faible"])]

            edite = st.data_editor(
                df, key=f"ed_{nom}_{filtre}", width="stretch", height=480,
                disabled=["Indicateur", "Page", "Extrait", "Commentaire / contrôles"],
                column_config={
                    "Indicateur": st.column_config.TextColumn(width="large"),
                    "Valeur": st.column_config.TextColumn(help="Modifiable : la cellule passera en « validé »"),
                    "Confiance": st.column_config.SelectboxColumn(options=pi.CONFIANCES, required=True),
                    "Page": st.column_config.NumberColumn(format="%d", width="small"),
                    "Extrait": st.column_config.TextColumn(width="large"),
                },
            )
            # report des corrections
            for ligne, row in edite.iterrows():
                r = res[ligne]
                ancienne = "" if r.valeur is None else str(r.valeur)
                nouvelle = (row["Valeur"] or "").strip()
                if nouvelle != ancienne:
                    r.valeur = nouvelle or None
                    r.confiance = "validé" if nouvelle else "non trouvé"
                    r.commentaire = (r.commentaire + " ; corrigé manuellement").strip(" ;")
                elif row["Confiance"] != r.confiance:
                    r.confiance = row["Confiance"]

            with st.expander("🔎 Voir une page du PDF pour vérifier"):
                pages_citees = sorted({r.page for r in res.values() if r.page})
                c1, c2 = st.columns([1, 4])
                p = c1.number_input("Page", 1, 999, pages_citees[0] if pages_citees else 1,
                                    key=f"p_{nom}")
                c1.caption("Pages citées : " + ", ".join(map(str, pages_citees)))
                c2.image(pi.image_page(ss.pdfs[nom], int(p)), width="stretch")

            if st.button("🗑️ Supprimer ce résultat", key=f"del_{nom}"):
                del ss.resultats[nom]
                ss.pdfs.pop(nom, None)
                st.rerun()

    # ------------------------------------------------------------ 3. export
    st.subheader("3. Télécharger le fichier Excel")
    xlsx = pi.ecrire_excel(modele_bytes, feuille, list(ss.resultats.items()))
    st.download_button(
        "⬇️ Télécharger l'Excel rempli", data=xlsx, type="primary",
        file_name=f"indicateurs_perfimmo_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Une colonne par rapport, cellules colorées selon la confiance, commentaire avec la page "
               "et l'extrait source. Feuilles « Détail extraction » (format long) et « Légende » ajoutées.")
