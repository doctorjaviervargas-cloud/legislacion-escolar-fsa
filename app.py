from __future__ import annotations

import hmac
import os
import re
import secrets
import shutil
import threading
import time
from collections import defaultdict, deque
from pathlib import Path
from typing import Any

from flask import (
    Flask, jsonify, redirect, render_template, request, session, url_for
)
from werkzeug.middleware.proxy_fix import ProxyFix

import ia_web_core as core

BASE_DIR = Path(__file__).resolve().parent
BUNDLED_NORMATIVA_DIR = BASE_DIR / "Normativa"
DATA_DIR = Path(os.getenv("APP_DATA_DIR", str(BASE_DIR / "data"))).resolve()
NORMATIVA_DIR = Path(os.getenv("NORMATIVA_DIR", str(DATA_DIR / "Normativa"))).resolve()
APP_TITLE = "Legislación Escolar de Formosa — Dr. Javier Vargas"
APP_VERSION = "Básica Online 4.3.2"
MAX_QUESTION_CHARS = int(os.getenv("MAX_QUESTION_CHARS", "2500"))
MAX_QUERIES_PER_MINUTE = int(os.getenv("MAX_QUERIES_PER_MINUTE", "20"))

INTERPRETATION_NOTICE = (
    "Esta interpretación tiene carácter informativo y orientativo y no sustituye "
    "el texto oficial de la norma ni el asesoramiento jurídico profesional aplicable "
    "al caso concreto. Para consultas, análisis de situaciones particulares o asistencia "
    "legal especializada en materia docente, puede contactar al Dr. Javier Vargas."
)

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "").strip() or secrets.token_hex(32)
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.getenv("COOKIE_SECURE", "1") == "1"
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

_lock = threading.Lock()
_rate_lock = threading.Lock()
_rate: dict[str, deque[float]] = defaultdict(deque)
_state: dict[str, Any] = {
    "units": [], "docs": [], "idf": {}, "contexts": [], "cdocs": [], "cidf": {},
    "files": [], "error": None, "loaded_at": None,
}


def _seed_normativa_if_needed() -> None:
    NORMATIVA_DIR.mkdir(parents=True, exist_ok=True)
    if any(NORMATIVA_DIR.glob("*.docx")):
        return
    if not BUNDLED_NORMATIVA_DIR.exists():
        return
    for src in BUNDLED_NORMATIVA_DIR.glob("*.docx"):
        if not src.name.startswith("~$"):
            shutil.copy2(src, NORMATIVA_DIR / src.name)


def load_normativa() -> None:
    _seed_normativa_if_needed()
    units, contexts, loaded = [], [], []
    error = None
    try:
        for path in sorted(NORMATIVA_DIR.glob("*.docx")):
            if path.name.startswith("~$"):
                continue
            blocks = core.extract_docx_blocks(path)
            units.extend(core.build_units_formatted(blocks, path.name))
            lines = core.extract_docx_lines(path)
            contexts.extend(core.build_contexts(lines, path.name))
            loaded.append(path.name)
        docs, idf = core.build_index(units)
        cdocs, cidf = core.build_context_index(contexts)
    except Exception as exc:
        docs, idf, cdocs, cidf = [], {}, [], {}
        error = f"No se pudo indexar la normativa: {exc}"
    with _lock:
        _state.update({
            "units": units, "docs": docs, "idf": idf,
            "contexts": contexts, "cdocs": cdocs, "cidf": cidf,
            "files": loaded, "error": error,
            "loaded_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        })


def snapshot() -> dict[str, Any]:
    with _lock:
        return dict(_state)


def find_decree_article(article: str, units: list[dict[str, Any]]) -> dict[str, Any] | None:
    article = str(article).strip().lower()
    for u in units:
        if (str(u.get("article", "")).strip().lower() == article
                and "1324/93" in u.get("norm_keys", [])
                and "931" not in u.get("norm_keys", [])):
            return u
    for u in units:
        if (str(u.get("article", "")).strip().lower() == article
                and "1324/93" in u.get("norm_keys", [])
                and (u.get("reg_text") or "").strip()):
            copy = dict(u)
            copy["text"] = copy.get("reg_text", "")
            copy["norm_labels"] = ["Decreto 1324/93"]
            copy["norm_keys"] = ["1324/93"]
            return copy
    return None


def unit_source(u: dict[str, Any]) -> dict[str, str]:
    return {
        "archivo": u.get("source", ""),
        "norma": " + ".join(u.get("norm_labels", [])),
        "articulo": str(u.get("article", "")),
        "texto": (u.get("text") or "").strip(),
    }


def safe_exact_response(question: str, found: list[dict[str, Any]], units: list[dict[str, Any]]) -> dict[str, Any]:
    article, norm_key = core.parse_query(question)
    if norm_key == "931" and article:
        law = next((u for u in units if str(u.get("article", "")).lower() == article and "931" in u.get("norm_keys", [])), found[0] if found else None)
        decree = find_decree_article(article, units)
        parts = [
            "MODO SEGURO — CONSULTA EXACTA", "",
            "LEY 931 — ESTATUTO DEL DOCENTE", f"Artículo {article}",
            (law.get("law_text") or law.get("text") or "").strip() if law else "No se encontró el texto de la Ley 931.",
            "", "DECRETO 1324/93 — REGLAMENTACIÓN", f"Artículo {article}",
            (decree.get("text") or "").strip() if decree else "No se detectó texto reglamentario para este artículo en la normativa cargada.",
        ]
        return {
            "mode":"exacta",
            "answer":"\n".join(parts),
            "sources":[],
            "ai_used":False
        }
    if found:
        u=found[0]
        return {
            "mode":"exacta",
            "answer":"MODO SEGURO — CONSULTA EXACTA\n\n" + f"{' + '.join(u.get('norm_labels', []))} — Artículo {u.get('article','')}\n\n" + (u.get("text") or "").strip(),
            "sources":[], "ai_used":False,
        }
    return {"mode":"sin coincidencias","answer":"No encontré respaldo suficiente en la normativa cargada.","sources":[],"ai_used":False}


def source_only_thematic(question: str, found: list[dict[str, Any]], context_hit: dict[str, Any] | None) -> dict[str, Any]:
    """
    Respuesta temática sin tarjetas de 'Fuente'.
    Toda la norma relacionada se incorpora una sola vez dentro de la respuesta.
    """
    if context_hit and context_hit.get("matched_phrase"):
        phrase=context_hit.get("matched_phrase","")
        explanation=core.sentence_around_phrase(context_hit.get("text",""), phrase)

        answer=[
            "MODO SEGURO — CONSULTA TEMÁTICA EXHAUSTIVA",
            "",
            f"Tema localizado: {phrase}",
            "",
            "Explicación encontrada en el Compendio:",
            explanation,
        ]

        if found:
            answer += [
                "",
                f"Normativa relacionada: {len(found)} referencia(s) encontrada(s)",
            ]

            for i,u in enumerate(found,1):
                label=" + ".join(u.get("norm_labels", []))
                article=u.get("article","")
                text=(u.get("text") or "").strip()

                answer += [
                    "",
                    f"{i}. {label} — Artículo {article}",
                    text,
                ]

        return {
            "mode":"temática exhaustiva",
            "answer":"\n".join(answer),
            "sources":[],
            "ai_used":False
        }

    if found:
        answer=[
            "RESULTADOS EN LA NORMATIVA CARGADA",
            "",
            f"Se localizaron {len(found)} referencia(s) relacionadas:"
        ]
        for i,u in enumerate(found,1):
            label=" + ".join(u.get("norm_labels", []))
            article=u.get("article","")
            text=(u.get("text") or "").strip()
            answer += [
                "",
                f"{i}. {label} — Artículo {article}",
                text,
            ]

        return {
            "mode":"temática",
            "answer":"\n".join(answer),
            "sources":[],
            "ai_used":False
        }

    return {
        "mode":"sin coincidencias",
        "answer":"No encontré respaldo suficiente en la normativa cargada.",
        "sources":[],
        "ai_used":False
    }



def _clean_for_interpretation(text: str) -> str:
    text = re.sub(r"\s+", " ", (text or "").strip())
    text = re.sub(r"^(?:ART(?:Í|I)?CULO\s+\d+(?:\s+(?:BIS|TER|QUATER|QUINQUIES))?\s*[º°o.]?\s*[-–—.:)]*\s*)", "", text, flags=re.I)
    return text.strip()


def _extractive_interpretation(question: str, found: list[dict[str, Any]], context_hit: dict[str, Any] | None) -> str:
    """Interpretación conservadora sin servicios externos.
    Selecciona los pasajes más vinculados con la consulta y los presenta como lectura
    orientativa, sin agregar requisitos, plazos o consecuencias que no estén en la fuente.
    """
    parts=[]
    if context_hit and context_hit.get("matched_phrase"):
        phrase=context_hit.get("matched_phrase","")
        expl=core.sentence_around_phrase(context_hit.get("text", ""), phrase)
        expl=_clean_for_interpretation(expl)
        if expl:
            parts.append(expl)

    qterms=set(core.toks(question))
    candidates=[]
    for u in found[:12]:
        txt=_clean_for_interpretation(u.get("text", ""))
        if not txt:
            continue
        # Divide en tramos breves y prioriza los que contienen términos consultados.
        sentences=[s.strip() for s in re.split(r"(?<=[.;:])\s+|\n+", txt) if s.strip()]
        for s in sentences:
            sterms=set(core.toks(s))
            score=len(qterms & sterms)
            if score or not qterms:
                candidates.append((score, len(s), s))

    candidates.sort(key=lambda x:(-x[0], x[1]))
    for _,_,s in candidates:
        if s not in parts:
            parts.append(s)
        if len(parts) >= 3:
            break

    if not parts and found:
        txt=_clean_for_interpretation(found[0].get("text", ""))
        if txt:
            parts=[txt[:900].strip()]

    if not parts:
        return "No es posible formular una interpretación orientativa con respaldo suficiente en la normativa recuperada."

    body=" ".join(parts)
    if len(body) > 1600:
        body=body[:1600].rsplit(" ",1)[0].rstrip(" ,;:") + "…"

    return (
        "Del contenido normativo recuperado se desprende, en términos orientativos, que "
        + body[0].lower() + body[1:] if body else
        "No es posible formular una interpretación orientativa con respaldo suficiente en la normativa recuperada."
    )


def interpretation_ai(question: str, found: list[dict[str, Any]], context_hit: dict[str, Any] | None) -> str | None:
    """Genera una interpretación estrictamente limitada a las normas recuperadas cuando
    hay una API de IA configurada. Si no existe, la aplicación usa el modo extractivo seguro.
    """
    api_key=os.getenv("OPENAI_API_KEY","").strip()
    if not api_key or not found:
        return None
    try:
        from openai import OpenAI
        model=os.getenv("OPENAI_MODEL","gpt-6-luna").strip() or "gpt-6-luna"
        client=OpenAI(api_key=api_key)
        source_blocks=[]
        for i,u in enumerate(found[:12],1):
            source_blocks.append(
                f"[NORMA {i}]\\nNorma: {' + '.join(u.get('norm_labels', []))}\\n"
                f"Artículo: {u.get('article','')}\\nTexto: {u.get('text','')}"
            )
        compendium=context_hit.get("text","") if context_hit else ""
        system=(
            "Sos un asistente de interpretación orientativa de legislación escolar de Formosa. "
            "Explicá en lenguaje claro únicamente lo que surge de los textos proporcionados. "
            "No inventes requisitos, excepciones, plazos, efectos, artículos, jurisprudencia ni hechos. "
            "No des asesoramiento sobre un caso concreto. Si hay ambigüedad, indicala. "
            "Redactá uno o dos párrafos breves y no repitas literalmente toda la norma."
        )
        user=(
            f"CONSULTA: {question}\\n\\nCONTEXTO DEL COMPENDIO:\\n{compendium}\\n\\n"
            "TEXTOS NORMATIVOS:\\n" + "\\n\\n".join(source_blocks)
        )
        response=client.responses.create(
            model=model,
            input=[{"role":"system","content":system},{"role":"user","content":user}]
        )
        text=(response.output_text or "").strip()
        return text or None
    except Exception:
        return None


def add_interpretation(result: dict[str, Any], question: str, found: list[dict[str, Any]], context_hit: dict[str, Any] | None=None) -> dict[str, Any]:
    if result.get("mode") in {"sin normativa", "sin coincidencias"} or not found:
        result["interpretation"]=""
        result["interpretation_ai"]=False
        result["professional_notice"]=""
        return result

    ai_text=interpretation_ai(question, found, context_hit)
    interpretation=ai_text or _extractive_interpretation(question, found, context_hit)
    result["interpretation"]=interpretation
    result["interpretation_ai"]=bool(ai_text)
    result["professional_notice"]=INTERPRETATION_NOTICE
    return result

def online_ai_answer(question: str, found: list[dict[str, Any]], context_hit: dict[str, Any] | None) -> str | None:
    api_key=os.getenv("OPENAI_API_KEY","").strip()
    if not api_key or not found:
        return None
    try:
        from openai import OpenAI
        model=os.getenv("OPENAI_MODEL","gpt-6-luna").strip() or "gpt-6-luna"
        client=OpenAI(api_key=api_key)
        source_blocks=[]
        for i,u in enumerate(found[:15],1):
            source_blocks.append(
                f"[FUENTE {i}]\nArchivo: {u.get('source','')}\nNorma: {' + '.join(u.get('norm_labels', []))}\nArtículo: {u.get('article','')}\nTexto literal:\n{u.get('text','')}"
            )
        compendium=context_hit.get("text","") if context_hit else ""
        system=(
            "Sos un asistente especializado en legislación escolar de la Provincia de Formosa, Argentina. "
            "Respondé exclusivamente con las fuentes proporcionadas. No inventes ni completes normas, artículos, "
            "plazos, excepciones, fechas o jurisprudencia. Si las fuentes no alcanzan, respondé exactamente: "
            "'No encontré respaldo suficiente en la normativa cargada.'. Diferenciá con claridad el texto normativo de la explicación."
        )
        user=f"PREGUNTA:\n{question}\n\nEXPLICACIÓN DEL COMPENDIO, SI EXISTE:\n{compendium}\n\nFUENTES NORMATIVAS:\n" + "\n\n".join(source_blocks)
        response=client.responses.create(model=model,input=[{"role":"system","content":system},{"role":"user","content":user}])
        return (response.output_text or "").strip() or None
    except Exception:
        return None


def answer_question(question: str) -> dict[str, Any]:
    st=snapshot(); units=st["units"]
    if not units:
        return {"mode":"sin normativa","answer":"No hay normativa indexada en el servidor.","sources":[],"ai_used":False,"interpretation":"","professional_notice":""}

    found,match_mode=core.retrieve(question,units,st["docs"],st["idf"])
    article,norm_key=core.parse_query(question)

    if article and norm_key and found:
        # Para Ley 931 agregamos también el artículo reglamentario a la base interpretativa.
        interp_found=list(found[:1])
        if norm_key == "931":
            decree=find_decree_article(article,units)
            if decree:
                interp_found.append(decree)
        result=safe_exact_response(question,found,units)
        return add_interpretation(result,question,interp_found,None)

    bridge_found,context_hit,bridge_mode=core.retrieve_thematic_bridge(
        question,st["contexts"],st["cdocs"],st["cidf"],units
    )
    if bridge_found:
        found=bridge_found

    # La respuesta normativa sigue siendo directa y verificable. La IA, si está
    # configurada, se reserva para el bloque separado de Interpretación orientativa.
    result=source_only_thematic(question,found,context_hit)
    return add_interpretation(result,question,found,context_hit)


def client_ip() -> str:
    return request.remote_addr or "desconocido"


def rate_limit_ok() -> bool:
    if MAX_QUERIES_PER_MINUTE <= 0:
        return True
    now=time.monotonic(); cutoff=now-60
    key=client_ip()
    with _rate_lock:
        q=_rate[key]
        while q and q[0] < cutoff: q.popleft()
        if len(q) >= MAX_QUERIES_PER_MINUTE:
            return False
        q.append(now)
        return True



def access_password() -> str:
    return os.getenv("ACCESS_PASSWORD","").strip()


def access_ok() -> bool:
    return not access_password() or bool(session.get("access_ok"))


@app.before_request
def enforce_optional_access():
    if request.endpoint in {"static","health","access_login"}:
        return None
    if request.path.startswith("/static/"):
        return None
    if not access_ok():
        if request.path.startswith("/api/"):
            return jsonify({"error":"Acceso no autorizado."}),401
        return redirect(url_for("access_login", next=request.path))
    return None


@app.after_request
def security_headers(resp):
    resp.headers["X-Content-Type-Options"]="nosniff"
    resp.headers["X-Frame-Options"]="SAMEORIGIN"
    resp.headers["Referrer-Policy"]="strict-origin-when-cross-origin"
    resp.headers["Permissions-Policy"]="camera=(), microphone=(), geolocation=()"
    if request.path.startswith("/api/"):
        resp.headers["Cache-Control"]="no-store"
    return resp


@app.get("/")
def home():
    st=snapshot()
    return render_template("index.html",app_title=APP_TITLE,version=APP_VERSION,
        ai_enabled=bool(os.getenv("OPENAI_API_KEY","").strip()),document_count=len(st["files"]),unit_count=len(st["units"]))


@app.post("/api/consultar")
def consultar():
    if not rate_limit_ok():
        return jsonify({"error":"Se alcanzó el límite de consultas por minuto. Intente nuevamente en unos instantes."}),429
    payload=request.get_json(silent=True) or {}
    question=str(payload.get("pregunta","")).strip()
    if not question:
        return jsonify({"error":"Escriba una consulta."}),400
    if len(question) > MAX_QUESTION_CHARS:
        return jsonify({"error":f"La consulta es demasiado extensa. Máximo: {MAX_QUESTION_CHARS} caracteres."}),400
    return jsonify(answer_question(question))


@app.get("/api/estado")
def estado():
    st=snapshot()
    return jsonify({"version":APP_VERSION,"edicion":"basica","normativa":"solo lectura","archivos":st["files"],"articulos_indexados":len(st["units"]),"ia_online":bool(os.getenv("OPENAI_API_KEY","").strip()),"cargado":st["loaded_at"],"error":st["error"]})


@app.route("/acceso", methods=["GET","POST"])
def access_login():
    if not access_password():
        return redirect(url_for("home"))
    error=None
    if request.method=="POST":
        supplied=request.form.get("password","")
        if hmac.compare_digest(supplied,access_password()):
            session["access_ok"]=True
            return redirect(request.args.get("next") or url_for("home"))
        error="Clave incorrecta."
    return render_template("login.html",title="Acceso",subtitle="Legislación Escolar de Formosa",error=error,button="Ingresar")



@app.get("/health")
def health():
    st=snapshot()
    return jsonify({
        "ok": True,
        "normativa_cargada": bool(st["units"]),
        "version": APP_VERSION,
        "articulos": len(st["units"]),
        "error": st["error"],
    }), 200


load_normativa()

if __name__ == "__main__":
    port=int(os.getenv("PORT","5000"))
    app.run(host="0.0.0.0",port=port,debug=False)
