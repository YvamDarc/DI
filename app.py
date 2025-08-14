# app_admin.py — Admin only (FEC -> questions tabulaires par sous-compte)
# Colonnes: N°, Date, Libellé, Question, Montant, Pièce, Statut, Sous-compte, Groupe
# Comptes: 401 (fournisseurs), 411 (clients), 47* (comptes d'attente)
# Texte adapté selon sens Débit/Crédit pour chaque catégorie
# Édition via st.data_editor, suppression, renumérotation, export JSON/Excel
# st.query_params + parsing dates robuste

import os
import json
from pathlib import Path
from typing import List, Dict, Any, Optional

import streamlit as st
import pandas as pd

# ----------------- Dossiers -----------------
BASE_DIR = Path(__file__).parent
EXPORTS_DIR = BASE_DIR / "exports"
EXPORTS_DIR.mkdir(exist_ok=True)

# ----------------- Utils -----------------
def _colmap(df: pd.DataFrame) -> dict:
    return {c.lower(): c for c in df.columns}

def _getcol(df: pd.DataFrame, *candidates: str) -> Optional[str]:
    cmap = _colmap(df)
    for cand in candidates:
        if cand.lower() in cmap:
            return cmap[cand.lower()]
    return None

def _starts_with(s: pd.Series, prefixes):
    if isinstance(prefixes, str):
        prefixes = (prefixes,)
    return s.astype(str).str.startswith(prefixes, na=False)

def _empty(s: pd.Series):
    return s.isna() | (s.astype(str).str.strip() == "")

def _to_naive_datetime(series: pd.Series) -> pd.Series:
    # Parse en tz-aware UTC puis enlève le tz -> naïf
    s = pd.to_datetime(series, errors="coerce", utc=True)
    return s.dt.tz_localize(None)

def _amount_series_abs(df: pd.DataFrame, cDebit: Optional[str], cCredit: Optional[str]) -> pd.Series:
    """Montant absolu (utile pour seuils lisibles)."""
    def to_num(s):
        if s.dtype == object:
            s = s.str.replace(",", ".", regex=False)
        return pd.to_numeric(s, errors="coerce").fillna(0.0)
    if cDebit and cDebit in df.columns and cCredit and cCredit in df.columns:
        d = to_num(df[cDebit])
        c = to_num(df[cCredit])
        return (d - c).abs()
    elif cDebit and cDebit in df.columns:
        return to_num(df[cDebit]).abs()
    elif cCredit and cCredit in df.columns:
        return to_num(df[cCredit]).abs()
    else:
        return pd.Series(0.0, index=df.index)

def _signed_amount_series(df: pd.DataFrame, cDebit: Optional[str], cCredit: Optional[str]) -> pd.Series:
    """Montant signé : >0 Débit, <0 Crédit."""
    def to_num(s):
        if s.dtype == object:
            s = s.str.replace(",", ".", regex=False)
        return pd.to_numeric(s, errors="coerce").fillna(0.0)
    d = to_num(df[cDebit]) if (cDebit and cDebit in df.columns) else pd.Series(0.0, index=df.index)
    c = to_num(df[cCredit]) if (cCredit and cCredit in df.columns) else pd.Series(0.0, index=df.index)
    return d - c

def _sens_from_signed(x: float) -> str:
    return "Débit" if (x or 0) > 0 else ("Crédit" if (x or 0) < 0 else "N/A")

def _fmt(s, missing=""):
    s = "" if s is None else str(s).strip()
    return s if s else missing

def _detect_group(compte: str) -> str:
    if compte.startswith("401"): return "Fournisseurs (401)"
    if compte.startswith("411"): return "Clients (411)"
    if compte.startswith("47"):  return "Comptes d'attente (47)"
    return "Autres"

def read_fec_autodetect(uploaded_file) -> pd.DataFrame:
    # 1) CSV auto-sep
    try:
        return pd.read_csv(uploaded_file, sep=None, engine="python", dtype=str)
    except Exception:
        try: uploaded_file.seek(0)
        except Exception: pass
    # 2) CSV ;
    try:
        return pd.read_csv(uploaded_file, sep=";", dtype=str)
    except Exception:
        try: uploaded_file.seek(0)
        except Exception: pass
    # 3) Excel
    return pd.read_excel(uploaded_file)

# ----------------- Texte de question adapté (compte & sens) -----------------
def _question_text_for(compte: str, sens: str, cas: str) -> str:
    """
    compte: '401...', '411...', '47...'
    sens: 'Débit' | 'Crédit' | 'N/A'
    cas: 'non_letre' | 'sans_piece' | 'doublon' | 'attente'
    """
    if compte.startswith("401"):
        if cas == "non_letre":
            return ("Fournisseur – poste non lettré ancien. "
                    + ("(Crédit 401 = facture reçue). " if sens=="Crédit" else "(Débit 401 = règlement/avoir/avance). ")
                    + "Merci de préciser le statut (litige, relance, avoir, plan de règlement) et joindre la pièce si disponible.")
        if cas == "sans_piece":
            return "Fournisseur – pièce absente : merci de fournir la facture/l’avoir/le relevé correspondant."
        if cas == "doublon":
            return ("Deux écritures similaires détectées (même date & montant). "
                    + ("Crédit 401 (probable facture) " if sens=="Crédit" else "Débit 401 (probable règlement/avoir) ")
                    + "— confirmer s’il s’agit d’un doublon ou préciser l’explication.")
    if compte.startswith("411"):
        if cas == "non_letre":
            return ("Client – poste non lettré ancien. "
                    + ("(Débit 411 = facture client). " if sens=="Débit" else "(Crédit 411 = encaissement/avoir). ")
                    + "Merci d’indiquer le statut de recouvrement et joindre la pièce si disponible.")
        if cas == "sans_piece":
            return "Client – pièce absente : merci de fournir la facture/avoir/relevé bancaire."
        if cas == "doublon":
            return ("Deux écritures similaires détectées (même date & montant). "
                    + ("Débit 411 (probable facture) " if sens=="Débit" else "Crédit 411 (probable encaissement/avoir) ")
                    + "— confirmer doublon ou expliquer.")
    if compte.startswith("47"):
        if cas in ("attente","sans_piece","non_letre"):
            base = "Compte d’attente 47 — "
            if sens == "Débit":
                base += "mouvement au débit. "
            elif sens == "Crédit":
                base += "mouvement au crédit. "
            return base + "Merci de préciser la nature réelle, le compte définitif et joindre le justificatif."
    # fallback générique
    if cas == "sans_piece":
        return "Pièce absente : merci de fournir le justificatif."
    if cas == "doublon":
        return "Deux écritures similaires détectées : merci de confirmer s’il s’agit d’un doublon ou d’expliquer."
    return "Merci de préciser la nature et fournir le justificatif."

# ----------------- Génération des lignes tabulaires -----------------
def generate_rows_tab(
    df: pd.DataFrame,
    aging_days: int,
    amount_threshold: float,
    max_rows: int
) -> pd.DataFrame:
    # mapping colonnes
    cCompteNum = _getcol(df, "CompteNum")
    cCompAuxNum = _getcol(df, "CompAuxNum")
    cEcrDate   = _getcol(df, "EcritureDate", "DateEcriture", "EcrDate")
    cPieceRef  = _getcol(df, "PieceRef", "ReferencePiece", "RefPiece")
    cDebit     = _getcol(df, "Debit")
    cCredit    = _getcol(df, "Credit")
    cLib       = _getcol(df, "EcritureLib", "LibelleEcriture")
    cLet       = _getcol(df, "EcritureLet", "Lettrage", "CodeLettrage")
    cDateLet   = _getcol(df, "DateLet", "LettrageDate")

    # diagnostic
    with st.expander("🔎 Colonnes détectées"):
        st.write({
            "CompteNum": cCompteNum, "CompAuxNum": cCompAuxNum,
            "EcritureDate": cEcrDate, "PieceRef": cPieceRef,
            "Debit": cDebit, "Credit": cCredit,
            "EcritureLib": cLib, "EcritureLet": cLet, "DateLet": cDateLet
        })

    # dates
    if cEcrDate:
        df[cEcrDate] = _to_naive_datetime(df[cEcrDate])
        dt_i64 = df[cEcrDate].view("int64")
        cutoff_i64 = (pd.Timestamp.utcnow().normalize() - pd.Timedelta(days=aging_days)).value
    else:
        dt_i64 = None
        cutoff_i64 = None

    # montants (absolu + signé)
    amt_abs = _amount_series_abs(df, cDebit, cCredit)
    amt_signed = _signed_amount_series(df, cDebit, cCredit)

    def sous_compte(row):
        aux = _fmt(row.get(cCompAuxNum, ""))
        return aux if aux else _fmt(row.get(cCompteNum, ""))

    rows: List[Dict[str, Any]] = []

    # 1) 401/411 non lettrés & anciens
    if cCompteNum:
        mask_tiers = _starts_with(df[cCompteNum], ("401","411"))
        if cLet or cDateLet:
            m_unlet = _empty(df[cLet]) if cLet else pd.Series(True, index=df.index)
            if cDateLet: m_unlet = m_unlet | df[cDateLet].isna()
        else:
            m_unlet = pd.Series(True, index=df.index)
        if cEcrDate and dt_i64 is not None:
            m_old = dt_i64.notna() & (dt_i64 <= cutoff_i64)
        else:
            m_old = pd.Series(True, index=df.index)
        m_amt = (amt_abs >= float(amount_threshold))

        cand = df[mask_tiers & m_unlet & m_old & m_amt].copy()
        for idx, r in cand.iterrows():
            comp = _fmt(r.get(cCompteNum, ""))
            sens = _sens_from_signed(float(amt_signed.loc[idx]))
            grp  = _detect_group(comp)
            sc   = sous_compte(r)
            d    = r[cEcrDate].date().isoformat() if cEcrDate and pd.notna(r[cEcrDate]) else ""
            lib  = _fmt(r.get(cLib, ""))
            piece= _fmt(r.get(cPieceRef, ""))
            m    = round(float(amt_abs.loc[idx] or 0.0),2)
            question = _question_text_for(comp, sens, cas="non_letre")
            rows.append({
                "Sous-compte": sc, "Groupe": grp, "Date": d, "Libellé": lib,
                "Montant": m, "Pièce": piece, "Question": question, "Statut": ""
            })
            if len(rows) >= max_rows: return pd.DataFrame(rows)

    # 2) 401/411 sans référence de pièce
    if cCompteNum and cPieceRef:
        mask_tiers = _starts_with(df[cCompteNum], ("401","411"))
        m_nopic = _empty(df[cPieceRef])
        m_amt = (amt_abs >= float(amount_threshold))
        miss = df[mask_tiers & m_nopic & m_amt].copy()
        for idx, r in miss.iterrows():
            comp = _fmt(r.get(cCompteNum, ""))
            sens = _sens_from_signed(float(amt_signed.loc[idx]))
            grp  = _detect_group(comp)
            sc   = sous_compte(r)
            d    = r[cEcrDate].date().isoformat() if cEcrDate and pd.notna(r[cEcrDate]) else ""
            lib  = _fmt(r.get(cLib, ""))
            m    = round(float(amt_abs.loc[idx] or 0.0),2)
            question = _question_text_for(comp, sens, cas="sans_piece")
            rows.append({
                "Sous-compte": sc, "Groupe": grp, "Date": d, "Libellé": lib,
                "Montant": m, "Pièce": "", "Question": question, "Statut": ""
            })
            if len(rows) >= max_rows: return pd.DataFrame(rows)

    # 3) Doublons règlements fournisseurs (401)
    if cCompteNum and cEcrDate:
        tmp = df[_starts_with(df[cCompteNum], "401")].copy()
        if not tmp.empty:
            tmp[cEcrDate] = _to_naive_datetime(tmp[cEcrDate])
            tmp["__date__"] = tmp[cEcrDate].dt.date
            s_signed = _signed_amount_series(tmp, _getcol(tmp,"Debit"), _getcol(tmp,"Credit"))
            tmp["__amt_signed__"] = s_signed
            tmp["__amt__"] = s_signed.abs()
            tmp = tmp[tmp["__amt__"] >= float(amount_threshold)]
            tcol = _getcol(tmp, "CompAuxNum") or cCompteNum
            if tcol in tmp.columns:
                grp_idx = tmp.groupby([tcol, "__date__", "__amt__"]).size().reset_index(name="n")
                keys = grp_idx[grp_idx["n"] >= 2][[tcol, "__date__", "__amt__"]]
                if not keys.empty:
                    merged = tmp.merge(keys, on=[tcol, "__date__", "__amt__"], how="inner")
                    for _, r in merged.iterrows():
                        comp = _fmt(r.get(cCompteNum,""))
                        sens = _sens_from_signed(float(r["__amt_signed__"]))
                        sc   = _fmt(r.get(tcol,""))
                        d    = r["__date__"].isoformat() if pd.notna(r["__date__"]) else ""
                        lib  = _fmt(r.get(cLib,""))
                        piece= _fmt(r.get(cPieceRef,""))
                        m    = round(float(r["__amt__"] or 0.0),2)
                        question = _question_text_for(comp, sens, cas="doublon")
                        rows.append({
                            "Sous-compte": sc, "Groupe": "Fournisseurs (401)", "Date": d, "Libellé": lib,
                            "Montant": m, "Pièce": piece, "Question": question, "Statut": ""
                        })
                        if len(rows) >= max_rows: return pd.DataFrame(rows)

    # 4) Comptes d’attente 47* (toutes écritures 47…)
    if cCompteNum:
        ca = df[_starts_with(df[cCompteNum], "47")].copy()
        if not ca.empty:
            # pré-calcul du sens pour tout 47*
            s_signed_ca = _signed_amount_series(ca, cDebit, cCredit)
            for idx, r in ca.iterrows():
                comp = _fmt(r.get(cCompteNum,""))
                sens = _sens_from_signed(float(s_signed_ca.loc[idx] if idx in s_signed_ca.index else 0.0))
                sc   = _fmt(r.get(cCompAuxNum,"")) or comp
                d    = r[cEcrDate].date().isoformat() if cEcrDate and pd.notna(r.get(cEcrDate, None)) else ""
                lib  = _fmt(r.get(cLib,""))
                piece= _fmt(r.get(cPieceRef,""))
                question = _question_text_for(comp, sens, cas="attente")
                rows.append({
                    "Sous-compte": sc, "Groupe": "Comptes d'attente (47)", "Date": d, "Libellé": lib,
                    "Montant": "", "Pièce": piece, "Question": question, "Statut": ""
                })
                if len(rows) >= max_rows: return pd.DataFrame(rows)

    return pd.DataFrame(rows)

def renumeroter(dfq: pd.DataFrame, par_sous_compte: bool) -> pd.DataFrame:
    dfq = dfq.copy()
    if dfq.empty:
        dfq["N°"] = []
        return dfq
    if par_sous_compte:
        dfq = dfq.sort_values(["Groupe","Sous-compte","Date","Libellé"], na_position="last")
        dfq["N°"] = (
            dfq.groupby(["Groupe","Sous-compte"])
               .cumcount() + 1
        )
    else:
        dfq = dfq.sort_values(["Groupe","Sous-compte","Date","Libellé"], na_position="last").reset_index(drop=True)
        dfq["N°"] = dfq.index + 1
    return dfq

# ----------------- UI -----------------
st.set_page_config(page_title="Préparation formulaire (Admin)", page_icon="🧾", layout="wide")
st.title("👩‍💼 Préparation du formulaire (comptable)")
st.caption("Importer un FEC → Générer des questions → Modifier/Supprimer → Organiser par sous-compte → Export JSON/Excel")

qp = st.query_params
titre = st.text_input("Titre du formulaire (pour l'export)", value=qp.get("title", "Formulaire client"))

st.subheader("1) Importer un FEC")
fec_file = st.file_uploader("Fichier FEC (CSV / TXT / Excel)", type=["csv","txt","xlsx","xls"])

st.subheader("2) Paramètres d’analyse")
c1, c2, c3, c4 = st.columns(4)
with c1:
    aging = st.number_input("Ancienneté (jours) pour 401/411 non lettrés", min_value=0, max_value=365, value=90, step=15)
with c2:
    seuil = st.number_input("Seuil de matérialité (montant min. €)", min_value=0.0, max_value=1_000_000.0, value=100.0, step=50.0, format="%.2f")
with c3:
    max_rows = st.number_input("Nombre max. de lignes générées", min_value=10, max_value=10000, value=2000, step=50)
with c4:
    renum_par_sc = st.checkbox("Renuméroter par sous-compte", value=True)

st.divider()

if "dfq" not in st.session_state:
    st.session_state["dfq"] = pd.DataFrame()

b1, b2 = st.columns([1,1])
with b1:
    if st.button("Analyser le FEC et générer"):
        if fec_file is None:
            st.warning("Veuillez importer un FEC.")
        else:
            try:
                df_fec = read_fec_autodetect(fec_file)
                dfq = generate_rows_tab(df_fec, aging_days=int(aging), amount_threshold=float(seuil), max_rows=int(max_rows))
                if dfq.empty:
                    st.info("Aucune ligne générée avec ces critères.")
                else:
                    dfq = renumeroter(dfq, par_sous_compte=renum_par_sc)
                    # ordre colonnes
                    cols = ["N°","Date","Libellé","Question","Montant","Pièce","Statut","Sous-compte","Groupe"]
                    for c in cols:
                        if c not in dfq.columns: dfq[c] = ""
                    st.session_state["dfq"] = dfq[cols]
                    st.success(f"{len(dfq)} ligne(s) prête(s). Vous pouvez éditer/supprimer ci-dessous.")
            except Exception as e:
                st.error(f"Erreur d'analyse du FEC: {e}")
with b2:
    if st.button("Vider"):
        st.session_state["dfq"] = pd.DataFrame()

st.subheader("3) Édition et organisation par sous-compte")

dfq = st.session_state["dfq"]
if not dfq.empty:
    # tri par Groupe/Sous-compte/Date
    dfq = dfq.sort_values(["Groupe","Sous-compte","Date","Libellé"], na_position="last")
    st.session_state["dfq"] = dfq

    # Affichage groupé par sous-compte
    for (grp, sc), df_sub in dfq.groupby(["Groupe","Sous-compte"], sort=False):
        with st.expander(f"{grp} — Sous-compte {sc}  ({len(df_sub)} lignes)", expanded=False):
            # Ajouter colonne suppression (locale à l’éditeur)
            df_edit = df_sub.copy()
            df_edit["Supprimer"] = False

            edited = st.data_editor(
                df_edit,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "N°": st.column_config.NumberColumn("N°", width="small"),
                    "Date": st.column_config.TextColumn("Date", help="AAAA-MM-JJ"),
                    "Libellé": st.column_config.TextColumn("Libellé"),
                    "Question": st.column_config.TextColumn("Question"),
                    "Montant": st.column_config.NumberColumn("Montant", step=0.01),
                    "Pièce": st.column_config.TextColumn("Pièce"),
                    "Statut": st.column_config.TextColumn("Statut", help="litige / relance / avoir / plan de règlement / ..."),
                    "Sous-compte": st.column_config.TextColumn("Sous-compte", width="small"),
                    "Groupe": st.column_config.TextColumn("Groupe", width="small"),
                    "Supprimer": st.column_config.CheckboxColumn("🗑️", help="Cocher pour supprimer"),
                },
                hide_index=True,
                key=f"editor_{grp}_{sc}"
            )

            # Appliquer modifications de ce sous-compte
            if st.button(f"Appliquer modifications — {grp} / {sc}", key=f"apply_{grp}_{sc}"):
                # remplacer les lignes correspondantes par edited (hors Supprimer cochés)
                keep_mask = ~((st.session_state["dfq"]["Groupe"] == grp) & (st.session_state["dfq"]["Sous-compte"] == sc))
                remain = st.session_state["dfq"][keep_mask]
                edited_clean = edited[edited["Supprimer"] != True].drop(columns=["Supprimer"], errors="ignore")
                st.session_state["dfq"] = pd.concat([remain, edited_clean], ignore_index=True)
                # renumérotation si demandé
                st.session_state["dfq"] = renumeroter(st.session_state["dfq"], par_sous_compte=renum_par_sc)
                st.success("Modifications appliquées.")

    st.divider()
    if st.button("Renuméroter maintenant"):
        st.session_state["dfq"] = renumeroter(st.session_state["dfq"], par_sous_compte=renum_par_sc)
        st.success("Renumérotation effectuée.")

    st.subheader("4) Export")
    colx, coly = st.columns(2)
    with colx:
        if st.button("Exporter en JSON"):
            out = EXPORTS_DIR / f"{titre.replace(' ','_')}.json"
            st.session_state["dfq"].to_json(out, orient="records", force_ascii=False, indent=2)
            with open(out, "rb") as fh:
                st.download_button("Télécharger le JSON", fh, file_name=out.name)
    with coly:
        if st.button("Exporter en Excel"):
            out = EXPORTS_DIR / f"{titre.replace(' ','_')}.xlsx"
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                # Feuille par groupe pour lisibilité
                for grp, df_grp in st.session_state["dfq"].groupby("Groupe"):
                    df_grp.to_excel(writer, sheet_name=grp[:31], index=False)
                # Feuille 'Tout'
                st.session_state["dfq"].to_excel(writer, sheet_name="Tout", index=False)
            with open(out, "rb") as fh:
                st.download_button("Télécharger l'Excel", fh, file_name=out.name)
else:
    st.info("Aucune ligne à afficher. Importez un FEC et cliquez sur « Analyser le FEC et générer ».") 
