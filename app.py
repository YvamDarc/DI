# app_admin.py — Admin only (FEC -> questions -> éditer -> exporter)
# Corrigé : dates homogénéisées (naïves) + comparaison int64
# UI : upload FEC, génération (401/411/471), édition, suppression, ordre, export JSON/Word
# Query params : st.query_params

import os
import json
from pathlib import Path
from typing import List, Dict, Any, Optional

import streamlit as st
import pandas as pd
from docx import Document

# ----------------- Dossiers -----------------
BASE_DIR = Path(__file__).parent
FORMS_DIR = BASE_DIR / "forms"     # export JSON
EXPORTS_DIR = BASE_DIR / "exports" # export Word
FORMS_DIR.mkdir(exist_ok=True)
EXPORTS_DIR.mkdir(exist_ok=True)

SUPPORTED_TYPES = ["oui_non", "texte", "checkbox_multi", "fichier"]

# ----------------- Utils génériques -----------------
def normalize_options(opt_str: str) -> List[str]:
    if not isinstance(opt_str, str):
        return []
    return [p.strip() for p in opt_str.split(";") if p.strip()]

def export_word(questions: List[Dict[str, Any]], title: str) -> Path:
    doc = Document()
    doc.add_heading(title or "Formulaire – Client", 0)
    for i, q in enumerate(questions, start=1):
        doc.add_paragraph(f"Q{i}. {q['question']}")
    out = EXPORTS_DIR / f"{(title or 'formulaire').replace(' ', '_')}.docx"
    doc.save(out)
    return out

def save_form_json(questions: List[Dict[str, Any]], filename: str) -> Path:
    out = FORMS_DIR / f"{(filename or 'formulaire').replace(' ', '_')}.json"
    out.write_text(json.dumps(questions, ensure_ascii=False, indent=2), encoding="utf-8")
    return out

# ----------------- FEC helpers -----------------
def _colmap(df: pd.DataFrame) -> dict:
    return {c.lower(): c for c in df.columns}

def _getcol(df: pd.DataFrame, *candidates: str) -> Optional[str]:
    cmap = _colmap(df)
    for cand in candidates:
        if cand.lower() in cmap:
            return cmap[cand.lower()]
    return None

def _starts_with_series(s: pd.Series, prefixes):
    if isinstance(prefixes, str):
        prefixes = (prefixes,)
    return s.astype(str).str.startswith(prefixes, na=False)

def _empty_series(s: pd.Series):
    return s.isna() | (s.astype(str).str.strip() == "")

def _to_naive_datetime(series: pd.Series) -> pd.Series:
    """Parse -> UTC aware -> drop tz -> naïf."""
    s = pd.to_datetime(series, errors="coerce", utc=True)   # tz-aware
    return s.dt.tz_localize(None)                           # drop tz => naïf

def _to_i64(series_dt_naive: pd.Series) -> pd.Series:
    """Convertit une série datetime64[ns] (naïve) en int64 (ns since epoch).
    NaT -> très petit int64 -> on gérera avec notna() à part.
    """
    return series_dt_naive.view("int64")

def _cutoff_i64(days: int) -> int:
    """Cutoff à J-<days>, au début de journée UTC, en int64 ns."""
    # Timestamp naïf (pas de tz) puis vue int64
    cutoff = (pd.Timestamp.utcnow().normalize() - pd.Timedelta(days=days))
    return cutoff.value  # int64 ns

def _make_amount_series(df: pd.DataFrame, cDebit: Optional[str], cCredit: Optional[str]) -> pd.Series:
    def to_num(s):
        if s.dtype == object:
            s = s.str.replace(",", ".", regex=False)
        return pd.to_numeric(s, errors="coerce")
    if cDebit and cDebit in df.columns and cCredit and cCredit in df.columns:
        d = to_num(df[cDebit]).fillna(0.0)
        c = to_num(df[cCredit]).fillna(0.0)
        return (d - c).abs()
    elif cDebit and cDebit in df.columns:
        return to_num(df[cDebit]).abs().fillna(0.0)
    elif cCredit and cCredit in df.columns:
        return to_num(df[cCredit]).abs().fillna(0.0)
    else:
        return pd.Series(0.0, index=df.index)

def read_fec_autodetect(uploaded_file) -> pd.DataFrame:
    # 1) CSV auto-sep
    try:
        return pd.read_csv(uploaded_file, sep=None, engine="python", dtype=str)
    except Exception:
        try:
            uploaded_file.seek(0)
        except Exception:
            pass
    # 2) CSV ;
    try:
        return pd.read_csv(uploaded_file, sep=";", dtype=str)
    except Exception:
        try:
            uploaded_file.seek(0)
        except Exception:
            pass
    # 3) Excel
    return pd.read_excel(uploaded_file)

def generate_questions_401_411_471(
    df: pd.DataFrame,
    aging_days: int = 90,
    top_n_per_bucket: int = 20,
    amount_threshold: float = 0.0
) -> List[Dict[str, Any]]:
    q: List[Dict[str, Any]] = []

    # mapping colonnes
    cCompteNum = _getcol(df, "CompteNum")
    cCompteLib = _getcol(df, "CompteLib")
    cCompAuxNum = _getcol(df, "CompAuxNum")
    cCompAuxLib = _getcol(df, "CompAuxLib")
    cEcrDate   = _getcol(df, "EcritureDate", "DateEcriture", "EcrDate")
    cPieceRef  = _getcol(df, "PieceRef", "ReferencePiece", "RefPiece")
    cDebit     = _getcol(df, "Debit")
    cCredit    = _getcol(df, "Credit")
    cLet       = _getcol(df, "EcritureLet", "Lettrage", "CodeLettrage")
    cDateLet   = _getcol(df, "DateLet", "LettrageDate")
    cLib       = _getcol(df, "EcritureLib", "LibelleEcriture")

    # --- DIAGNOSTIC rapide ---
    with st.expander("🔎 Diagnostic colonnes détectées"):
        st.write({
            "CompteNum": cCompteNum, "CompteLib": cCompteLib,
            "CompAuxNum": cCompAuxNum, "CompAuxLib": cCompAuxLib,
            "EcritureDate": cEcrDate, "PieceRef": cPieceRef,
            "Debit": cDebit, "Credit": cCredit,
            "EcritureLet": cLet, "DateLet": cDateLet,
            "EcritureLib": cLib
        })

    # Dates & montants
    if cEcrDate:
        # Série datetime naïve + int64
        dt_naive = _to_naive_datetime(df[cEcrDate])
        df[cEcrDate] = dt_naive                          # garde la version datetime naïve
        dt_i64 = _to_i64(dt_naive)                       # version int64 pour comparaisons
        cutoff_i64 = _cutoff_i64(aging_days)
    else:
        dt_i64 = None
        cutoff_i64 = None

    amt_series = _make_amount_series(df, cDebit, cCredit)

    def tier_id(row: pd.Series):
        aux = str(row.get(cCompAuxNum, "") or "").strip() if cCompAuxNum else ""
        auxlib = str(row.get(cCompAuxLib, "") or "").strip() if cCompAuxLib else ""
        if aux:
            return aux, auxlib
        compte = str(row.get(cCompteNum, "") or "").strip() if cCompteNum else ""
        complib = str(row.get(cCompteLib, "") or "").strip() if cCompteLib else ""
        return compte, complib

    # 1) 401/411 non lettrés anciens
    if cCompteNum:
        mask_tiers = _starts_with_series(df[cCompteNum], ("401","411"))
        if cLet or cDateLet:
            m_unlet = _empty_series(df[cLet]) if cLet else pd.Series(True, index=df.index)
            if cDateLet:
                m_unlet = m_unlet | df[cDateLet].isna()
        else:
            m_unlet = pd.Series(True, index=df.index)
        # ancienneté via int64 (robuste)
        if cEcrDate and dt_i64 is not None:
            m_date_ok = dt_i64.notna() & (dt_i64 <= cutoff_i64)
        else:
            m_date_ok = pd.Series(True, index=df.index)
        m_amt = (amt_series >= float(amount_threshold))

        cand = df[mask_tiers & m_unlet & m_date_ok & m_amt].copy()
        rows = []
        for idx, r in cand.iterrows():
            t_id, t_lib = tier_id(r)
            amt = float(amt_series.loc[idx] or 0.0)
            dt = r[cEcrDate] if cEcrDate else None
            rows.append({"tier": t_id or "(inconnu)", "lib": t_lib, "amount": amt, "date": dt})

        if rows:
            tb = pd.DataFrame(rows)
            agg = tb.groupby(["tier","lib"], dropna=False).agg(
                n=("amount","size"),
                total=("amount","sum"),
                oldest=("date","min")
            ).reset_index()
            agg = agg.sort_values("total", ascending=False).head(top_n_per_bucket)
            for _, r in agg.iterrows():
                t = r["tier"]
                lib = r["lib"] or ""
                tot = round(float(r["total"] or 0), 2)
                n = int(r["n"])
                oldest = r["oldest"]
                oldest_str = oldest.date().isoformat() if pd.notna(oldest) else "date inconnue"
                who = "client" if str(t).startswith("411") else "fournisseur"
                label = (
                    f"Postes non lettrés (> {aging_days} j) sur {who} {t} "
                    f"{('('+lib+')' if lib else '')}: {n} écriture(s), total ~{tot}€. "
                    f"Ancienneté depuis {oldest_str}. Indiquez le statut (litige, relance, avoir, plan de règlement)."
                )
                q.append({"type":"texte","question":label})
                q.append({"type":"fichier","question":f"Joindre justificatifs (factures/relevés) pour {t} {('('+lib+')' if lib else '')}."})

    # 2) 401/411 sans référence de pièce
    if cCompteNum and cPieceRef:
        mask_tiers = _starts_with_series(df[cCompteNum], ("401","411"))
        m_nopic = _empty_series(df[cPieceRef])
        m_amt = (amt_series >= float(amount_threshold))
        miss = df[mask_tiers & m_nopic & m_amt].copy()
        if len(miss):
            rows = []
            for idx, r in miss.iterrows():
                t_id, t_lib = tier_id(r)
                amt = float(amt_series.loc[idx] or 0.0)
                rows.append({"tier": t_id or "(inconnu)","lib": t_lib, "amount": amt})
            tb = pd.DataFrame(rows)
            agg = tb.groupby(["tier","lib"]).agg(n=("amount","size"), total=("amount","sum")).reset_index()
            agg = agg.sort_values("total", ascending=False).head(top_n_per_bucket)
            for _, r in agg.iterrows():
                t, lib, n, tot = r["tier"], r["lib"] or "", int(r["n"]), round(float(r["total"] or 0),2)
                q.append({"type":"texte","question":f"{n} écriture(s) sur {t} {('('+lib+')' if lib else '')} sans référence de pièce. Précisez la référence (ou raison)."})
                q.append({"type":"fichier","question":f"Joindre les pièces manquantes pour {t} {('('+lib+')' if lib else '')}."})

    # 3) Doublons de règlements fournisseurs
    if cCompteNum and cEcrDate:
        tmp = df[_starts_with_series(df[cCompteNum], "401")].copy()
        if not tmp.empty:
            tmp_dt_naive = _to_naive_datetime(tmp[cEcrDate])
            tmp_dt_i64 = _to_i64(tmp_dt_naive)
            tmp["__date__"] = pd.to_datetime(tmp_dt_naive, errors="coerce").dt.date
            amounts = _make_amount_series(tmp, _getcol(tmp, "Debit"), _getcol(tmp, "Credit"))
            tmp["__amt__"] = amounts
            tmp = tmp[tmp["__amt__"] >= float(amount_threshold)]
            tcol = _getcol(tmp, "CompAuxNum") or cCompteNum
            if tcol in tmp.columns:
                grp = tmp.groupby([tcol, "__date__", "__amt__"]).size().reset_index(name="n")
                dup = grp[grp["n"] >= 2].head(top_n_per_bucket)
                for _, r in dup.iterrows():
                    t = r[tcol]
                    date = r["__date__"]
                    amt = round(float(r["__amt__"] or 0),2)
                    q.append({"type":"oui_non","question":f"Deux règlements identiques détectés pour le fournisseur {t} le {date} pour ~{amt}€. Confirmez s'il s'agit d'un doublon ?"})
                    q.append({"type":"fichier","question":f"Joindre justificatif (relevé/annulation/avoir) concernant le doublon {t} {date} {amt}€."})

    # 4) Comptes d'attente 471
    if cCompteNum:
        ca = df[_starts_with_series(df[cCompteNum], "471")].copy()
        if len(ca):
            cLib = _getcol(ca, "EcritureLib", "LibelleEcriture")
            if cLib and (cLib in ca.columns):
                grp = ca.groupby(cLib).size().reset_index(name="n").sort_values("n", ascending=False).head(top_n_per_bucket)
                for _, r in grp.iterrows():
                    lib = str(r[cLib])[:120]
                    q.append({"type":"texte","question":f"Écritures en 471 détectées (ex: '{lib}'). Précisez la nature réelle et le compte définitif."})
                    q.append({"type":"fichier","question":f"Joindre justificatif pour l'écriture en 471 (ex: '{lib}')."})
            else:
                q.append({"type":"texte","question":"Écritures en 471 détectées. Précisez la nature réelle et la régularisation prévue."})
                q.append({"type":"fichier","question":"Joindre justificatifs pour les 471."})

    return q

# ----------------- UI ADMIN -----------------
st.set_page_config(page_title="Préparation formulaire (Admin)", page_icon="🧾", layout="centered")
st.title("👩‍💼 Préparation du formulaire (comptable)")
st.caption("Importer un FEC, générer des questions (401/411/471), modifier/supprimer/réordonner, puis exporter le formulaire.")

qp = st.query_params
form_title = st.text_input("Titre du formulaire (pour l'export)", value=qp.get("title", "Formulaire client"))

st.subheader("1) Importer un FEC")
fec_file = st.file_uploader("Fichier FEC (CSV / TXT / Excel)", type=["csv", "txt", "xlsx", "xls"])

st.subheader("2) Paramètres d'analyse (401/411/471)")
colA, colB, colC = st.columns(3)
with colA:
    aging = st.number_input("Ancienneté (jours) pour 401/411 non lettrés", min_value=30, max_value=365, value=90, step=15)
with colB:
    amount_thresh = st.number_input("Seuil de matérialité (montant min., €)", min_value=0.0, max_value=1_000_000.0, value=100.0, step=50.0, format="%.2f")
with colC:
    topn = st.number_input("Limite de questions par catégorie", min_value=5, max_value=200, value=20, step=5)

st.divider()

if "questions" not in st.session_state:
    st.session_state["questions"] = []

c1, c2 = st.columns([1,1])
with c1:
    if st.button("Analyser le FEC et générer des questions"):
        if fec_file is None:
            st.warning("Veuillez importer un FEC.")
        else:
            try:
                df_fec = read_fec_autodetect(fec_file)
                qs = generate_questions_401_411_471(
                    df_fec,
                    aging_days=int(aging),
                    top_n_per_bucket=int(topn),
                    amount_threshold=float(amount_thresh)
                )
                if not qs:
                    st.info("Aucune question pertinente détectée (selon vos paramètres).")
                else:
                    st.session_state["questions"] = qs
                    st.success(f"{len(qs)} question(s) générée(s). Vous pouvez maintenant modifier/supprimer/réordonner.")
            except Exception as e:
                st.error(f"Erreur d'analyse du FEC: {e}")
with c2:
    if st.button("Vider la liste"):
        st.session_state["questions"] = []

st.subheader("3) Édition du formulaire (modifier / supprimer / réordonner)")
qs = st.session_state["questions"]

with st.expander("➕ Ajouter une question manuellement"):
    a1, a2 = st.columns([2,1])
    with a1:
        new_label = st.text_input("Texte de la question", key="new_label_admin")
    with a2:
        new_type = st.selectbox("Type", SUPPORTED_TYPES, key="new_type_admin")
    new_opts_str = st.text_input("Options (séparées par ';') pour checkbox_multi", key="new_opts_admin")
    if st.button("Ajouter la question manuelle"):
        q = {"type": new_type, "question": new_label.strip() if new_label else "", "options": normalize_options(new_opts_str)}
        if q["question"]:
            qs.append(q)
            st.session_state["questions"] = qs
            st.success("Question ajoutée.")
        else:
            st.warning("Le texte de la question est vide.")

to_delete = []
reordered = []
for i, q in enumerate(qs):
    st.markdown(f"**Q{i+1}**")
    b1, b2, b3, b4 = st.columns([4,2,1,1])
    with b1:
        q["question"] = st.text_input("Intitulé", value=q["question"], key=f"label_{i}")
    with b2:
        q["type"] = st.selectbox("Type", SUPPORTED_TYPES, index=SUPPORTED_TYPES.index(q["type"]), key=f"type_{i}")
    with b3:
        pos = st.number_input("Ordre", min_value=1, max_value=len(qs), value=i+1, key=f"pos_{i}")
    with b4:
        if st.button("🗑️", key=f"del_{i}", help="Supprimer cette question"):
            to_delete.append(i)

    if q["type"] == "checkbox_multi":
        opt_str = "; ".join(q.get("options", []))
        new_opt_str = st.text_input("Options (; séparées)", value=opt_str, key=f"opts_{i}")
        q["options"] = normalize_options(new_opt_str)

    reordered.append((pos, q))

if to_delete:
    for idx in sorted(to_delete, reverse=True):
        qs.pop(idx)
    st.session_state["questions"] = qs
    st.rerun()

if st.button("Appliquer l'ordre"):
    reordered.sort(key=lambda x: x[0])
    qs = [q for _, q in reordered]
    st.session_state["questions"] = qs
    st.success("Ordre mis à jour.")

st.subheader("4) Exporter le formulaire")
e1, e2 = st.columns(2)
with e1:
    if st.button("Exporter en JSON"):
        p = save_form_json(st.session_state["questions"], filename=qp.get("title", "Formulaire client"))
        with open(p, "rb") as fh:
            st.download_button("Télécharger le JSON", fh, file_name=p.name)
with e2:
    if st.button("Exporter en Word"):
        if not st.session_state["questions"]:
            st.warning("Pas de questions à exporter.")
        else:
            p = export_word(st.session_state["questions"], title=qp.get("title", "Formulaire client"))
            with open(p, "rb") as fh:
                st.download_button("Télécharger le .docx", fh, file_name=p.name)
