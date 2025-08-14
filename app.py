# app_admin.py — Admin only
# FEC -> Génération de questions "par écriture" (401/411/471)
# Edition (modifier/supprimer/ordre) + Export JSON/Word
# st.query_params + dates homogénéisées, comparaisons robustes

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
    return [p.strip() for p in str(opt_str).split(";") if p.strip()]

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
    """datetime64[ns] -> int64 (ns since epoch)."""
    return series_dt_naive.view("int64")

def _cutoff_i64(days: int) -> int:
    """Cutoff à J-<days>, début de journée UTC, en int64 ns."""
    return (pd.Timestamp.utcnow().normalize() - pd.Timedelta(days=days)).value

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

# ----------------- Génération "par écriture" -----------------
def _fmt(val, missing="(inconnu)"):
    s = str(val) if val is not None else ""
    s = s.strip()
    return s if s else missing

def _fmt_money(x: float) -> str:
    try:
        return f"{round(float(x or 0.0), 2)}"
    except Exception:
        return "0.00"

def _q_for_entry(date, lib, amt, piece, compte, suffix="Pouvez-vous nous fournir la pièce manquante ou préciser la nature de cette écriture ?"):
    date_str = date.date().isoformat() if pd.notna(date) else "(date inconnue)"
    lib_s = _fmt(lib, "(libellé manquant)")
    piece_s = _fmt(piece, "(pièce absente)")
    compte_s = _fmt(compte, "(compte inconnu)")
    amt_s = _fmt_money(amt)
    text = (
        f"Écriture du {date_str} — \"{lib_s}\" — montant : {amt_s} € — pièce : {piece_s} — compte : {compte_s}. "
        f"{suffix}"
    )
    upload = f"Joindre justificatif pour l'écriture du {date_str} (\"{lib_s}\", {amt_s} €, compte {compte_s})"
    return text, upload

def generate_questions_401_411_471_per_entry(
    df: pd.DataFrame,
    aging_days: int = 90,
    amount_threshold: float = 0.0,
    max_questions: int = 300
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

    # Diagnostic (utile si entêtes exotiques)
    with st.expander("🔎 Colonnes détectées"):
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
        dt_naive = _to_naive_datetime(df[cEcrDate])   # datetime naïf
        df[cEcrDate] = dt_naive
        dt_i64 = _to_i64(dt_naive)
        cutoff_i64 = _cutoff_i64(aging_days)
    else:
        dt_i64 = None
        cutoff_i64 = None

    amt_series = _make_amount_series(df, cDebit, cCredit)

    # ---------- 1) 401/411 non lettrés ET anciens ----------
    if cCompteNum:
        mask_tiers = _starts_with_series(df[cCompteNum], ("401", "411"))
        if cLet or cDateLet:
            m_unlet = _empty_series(df[cLet]) if cLet else pd.Series(True, index=df.index)
            if cDateLet:
                m_unlet = m_unlet | df[cDateLet].isna()
        else:
            m_unlet = pd.Series(True, index=df.index)
        if cEcrDate and dt_i64 is not None:
            m_old = dt_i64.notna() & (dt_i64 <= cutoff_i64)
        else:
            m_old = pd.Series(True, index=df.index)
        m_amt = (amt_series >= float(amount_threshold))

        cand = df[mask_tiers & m_unlet & m_old & m_amt].copy()
        for idx, r in cand.iterrows():
            text, upload = _q_for_entry(
                date=r[cEcrDate] if cEcrDate else None,
                lib=r.get(cLib, ""),
                amt=amt_series.loc[idx],
                piece=r.get(cPieceRef, ""),
                compte=r.get(cCompteNum, ""),
                suffix="Merci de préciser le statut (litige, relance, avoir, plan de règlement) et de joindre la pièce si disponible."
            )
            q.append({"type": "texte", "question": text})
            q.append({"type": "fichier", "question": upload})
            if len(q) >= max_questions:
                return q

    # ---------- 2) 401/411 sans référence de pièce ----------
    if cCompteNum and cPieceRef:
        mask_tiers = _starts_with_series(df[cCompteNum], ("401", "411"))
        m_nopic = _empty_series(df[cPieceRef])
        m_amt = (amt_series >= float(amount_threshold))
        miss = df[mask_tiers & m_nopic & m_amt].copy()
        for idx, r in miss.iterrows():
            text, upload = _q_for_entry(
                date=r[cEcrDate] if cEcrDate else None,
                lib=r.get(cLib, ""),
                amt=amt_series.loc[idx],
                piece="(pièce absente)",
                compte=r.get(cCompteNum, ""),
                suffix="Pouvez-vous nous fournir la pièce manquante (facture/avoir/relevé) ?"
            )
            q.append({"type": "texte", "question": text})
            q.append({"type": "fichier", "question": upload})
            if len(q) >= max_questions:
                return q

    # ---------- 3) Doublons de règlements fournisseurs ----------
    if cCompteNum and cEcrDate:
        tmp = df[_starts_with_series(df[cCompteNum], "401")].copy()
        if not tmp.empty:
            tmp_dt = _to_naive_datetime(tmp[cEcrDate])
            tmp["__date__"] = tmp_dt.dt.date
            amounts = _make_amount_series(tmp, _getcol(tmp, "Debit"), _getcol(tmp, "Credit"))
            tmp["__amt__"] = amounts
            tmp = tmp[tmp["__amt__"] >= float(amount_threshold)]
            tcol = _getcol(tmp, "CompAuxNum") or cCompteNum
            if tcol in tmp.columns:
                grp = tmp.groupby([tcol, "__date__", "__amt__"]).size().reset_index(name="n")
                dup_keys = grp[grp["n"] >= 2][[tcol, "__date__", "__amt__"]]
                if not dup_keys.empty:
                    merged = tmp.merge(dup_keys, on=[tcol, "__date__", "__amt__"], how="inner")
                    for _, r in merged.iterrows():
                        text, upload = _q_for_entry(
                            date=r[cEcrDate] if cEcrDate in r else None,
                            lib=r.get(cLib, ""),
                            amt=r["__amt__"],
                            piece=r.get(cPieceRef, ""),
                            compte=r.get(cCompteNum, ""),
                            suffix="Deux écritures similaires détectées ce jour-là pour le même montant. Confirmez s'il s'agit d'un doublon ou fournissez l'explication (annulation/avoir)."
                        )
                        q.append({"type": "oui_non", "question": "S'agit-il d'un doublon ?"})
                        q.append({"type": "texte", "question": text})
                        q.append({"type": "fichier", "question": upload})
                        if len(q) >= max_questions:
                            return q

    # ---------- 4) Comptes d'attente 471 ----------
    if cCompteNum:
        ca = df[_starts_with_series(df[cCompteNum], "471")].copy()
        for _, r in ca.iterrows():
            text, upload = _q_for_entry(
                date=r[cEcrDate] if cEcrDate else None,
                lib=r.get(cLib, ""),
                amt=None,
                piece=r.get(cPieceRef, ""),
                compte=r.get(cCompteNum, ""),
                suffix="Écriture en compte d'attente. Merci de préciser la nature réelle et le compte définitif, et de joindre tout justificatif utile."
            )
            q.append({"type": "texte", "question": text})
            q.append({"type": "fichier", "question": upload})
            if len(q) >= max_questions:
                return q

    return q

# ----------------- UI ADMIN -----------------
st.set_page_config(page_title="Préparation formulaire (Admin)", page_icon="🧾", layout="centered")
st.title("👩‍💼 Préparation du formulaire (comptable)")
st.caption("Importer un FEC, générer des questions par écriture (401/411/471), modifier/supprimer/réordonner, puis exporter le formulaire.")

qp = st.query_params
form_title = st.text_input("Titre du formulaire (pour l'export)", value=qp.get("title", "Formulaire client"))

st.subheader("1) Importer un FEC")
fec_file = st.file_uploader("Fichier FEC (CSV / TXT / Excel)", type=["csv", "txt", "xlsx", "xls"])

st.subheader("2) Paramètres d'analyse")
colA, colB, colC = st.columns(3)
with colA:
    aging = st.number_input("Ancienneté (jours) pour 401/411 non lettrés", min_value=0, max_value=365, value=90, step=15)
with colB:
    amount_thresh = st.number_input("Seuil de matérialité (montant min., €)", min_value=0.0, max_value=1_000_000.0, value=100.0, step=50.0, format="%.2f")
with colC:
    max_q = st.number_input("Limite totale de questions générées", min_value=10, max_value=2000, value=300, step=10)

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
                qs = generate_questions_401_411_471_per_entry(
                    df_fec,
                    aging_days=int(aging),
                    amount_threshold=float(amount_thresh),
                    max_questions=int(max_q)
                )
                if not qs:
                    st.info("Aucune question générée avec ces critères.")
                else:
                    st.session_state["questions"] = qs
                    st.success(f"{len(qs)} question(s) générée(s). Vous pouvez modifier/supprimer/réordonner ci-dessous.")
            except Exception as e:
                st.error(f"Erreur d'analyse du FEC: {e}")
with c2:
    if st.button("Vider la liste"):
        st.session_state["questions"] = []

# 3) Édition
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
        q["type"] = st.selectbox("Type", SUPPORTED_TYPES, index=SUPPORTED_TYPES.index(q.get("type","texte")), key=f"type_{i}")
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

# 4) Export
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
