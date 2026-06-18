import os, uuid, warnings
from datetime import datetime
from pathlib import Path
from io import BytesIO
from functools import wraps

warnings.filterwarnings("ignore")

from flask import (Flask, render_template, jsonify, request,
                   redirect, url_for, session, send_file)
import pandas as pd
import sqlite3

BASE_DIR   = Path(__file__).parent
DB_PATH    = BASE_DIR / "faturamento.db"
FAT_PATH   = Path("/Users/leonardocarmo/Library/CloudStorage/OneDrive-Pessoal/Energy/FATURAMENTO.xlsx")
CTRL_PATH  = BASE_DIR / "Controle_Medicoes.xlsx"

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.secret_key = "energy-fat-2026"

USERS = {
    "admin":  {"password": "admin123",   "role": "admin"},
    "energy": {"password": "energy2026", "role": "user"},
}

# ── DB ────────────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS medicoes (
                id TEXT PRIMARY KEY,
                empresa TEXT, gestor TEXT,
                contrato_num TEXT, contrato_nome TEXT,
                obra TEXT, cod TEXT, comp TEXT,
                provisao REAL DEFAULT 0,
                medicao REAL,
                pedido TEXT, nf TEXT, venc_nf TEXT,
                retencao REAL, impostos REAL,
                status TEXT DEFAULT 'previsto',
                obs TEXT,
                created_at TEXT, updated_at TEXT,
                delete_requested INTEGER DEFAULT 0,
                delete_requested_by TEXT,
                delete_requested_at TEXT
            );
            CREATE TABLE IF NOT EXISTS contratos (
                num TEXT PRIMARY KEY,
                nome TEXT,
                saldo REAL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS delete_requests (
                id TEXT PRIMARY KEY,
                medicao_id TEXT,
                requested_by TEXT, requested_at TEXT,
                obra TEXT, contrato_num TEXT, contrato_nome TEXT, comp TEXT
            );
        """)
        count = conn.execute("SELECT COUNT(*) FROM medicoes").fetchone()[0]
        if count == 0 and FAT_PATH.exists():
            _seed_from_xlsx(conn)

def _seed_from_xlsx(conn):
    df = pd.read_excel(FAT_PATH, sheet_name="index", header=1, usecols="B:S")
    df.columns = [
        "empresa", "contrato_num", "contrato_nome", "gestor",
        "obra", "cod", "comp", "municipio",
        "provisao", "valor_bruto_fat", "pedido", "nf",
        "protocolo", "valor_bruto", "retencao", "impostos",
        "valor_liquido", "data_recebimento",
    ]
    df = df[df["empresa"].notna()].copy()
    df["contrato_num"] = df["contrato_num"].astype(str).str.strip()
    df["comp"] = pd.to_datetime(df["comp"], errors="coerce").dt.strftime("%Y-%m")
    df["provisao"] = pd.to_numeric(df["provisao"], errors="coerce").fillna(0)
    df["_med"] = pd.to_numeric(df["valor_bruto_fat"], errors="coerce")

    def _status(row):
        nf  = str(row.get("nf",  "") or "").strip()
        ped = str(row.get("pedido", "") or "").strip()
        med = row.get("_med")
        if nf  and nf  != "nan": return "nf_emitida"
        if ped and ped != "nan": return "aprovado"
        if pd.notna(med) and med > 0: return "medicao"
        return "previsto"

    now = datetime.now().isoformat()
    for _, r in df.iterrows():
        ped = str(r.get("pedido", "") or "").strip()
        nf  = str(r.get("nf",    "") or "").strip()
        med = r.get("_med")
        conn.execute("""
            INSERT OR IGNORE INTO medicoes
            (id,empresa,gestor,contrato_num,contrato_nome,obra,cod,comp,
             provisao,medicao,pedido,nf,retencao,impostos,status,created_at,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            str(uuid.uuid4()),
            r.get("empresa"), r.get("gestor"),
            r.get("contrato_num"), r.get("contrato_nome"),
            r.get("obra"), r.get("cod"), r.get("comp"),
            float(r.get("provisao", 0)),
            float(med) if pd.notna(med) else None,
            ped if ped and ped != "nan" else None,
            nf  if nf  and nf  != "nan" else None,
            float(r["retencao"]) if pd.notna(r.get("retencao")) else None,
            float(r["impostos"]) if pd.notna(r.get("impostos")) else None,
            _status(r), now, now,
        ))
    # seed contracts
    for cnum in df["contrato_num"].dropna().unique():
        nome = df[df["contrato_num"] == cnum]["contrato_nome"].iloc[0]
        conn.execute("INSERT OR IGNORE INTO contratos(num,nome,saldo) VALUES(?,?,0)",
                     (str(cnum), str(nome)))


def _read_controle():
    """Read Controle_Medicoes.xlsx live (updated by the email automation)."""
    if not CTRL_PATH.exists():
        return []
    try:
        df = pd.read_excel(CTRL_PATH, sheet_name="Medições")
        df.columns = [
            "data_recebimento", "n_folha", "n_contrato",
            "periodo_inicio", "periodo_fim",
            "municipio", "fornecedor", "valor_total", "arquivo", "status",
        ]
        df = df[df["n_folha"].notna()].copy()
        df["n_contrato"]  = df["n_contrato"].astype(str).str.strip()
        df["n_folha"]     = df["n_folha"].astype(str).str.strip()
        df["valor_total"] = pd.to_numeric(df["valor_total"], errors="coerce").fillna(0)
        df["periodo_inicio"] = pd.to_datetime(df["periodo_inicio"], format="%d/%m/%Y", errors="coerce")
        df["periodo"] = df["periodo_inicio"].dt.strftime("%Y-%m")
        result = []
        for _, row in df.iterrows():
            result.append({
                "n_folha":          str(row["n_folha"]),
                "n_contrato":       str(row["n_contrato"]),
                "periodo":          str(row["periodo"] or ""),
                "municipio":        str(row.get("municipio") or ""),
                "fornecedor":       str(row.get("fornecedor") or ""),
                "valor_total":      float(row["valor_total"]),
                "arquivo":          str(row.get("arquivo") or ""),
                "data_recebimento": str(row.get("data_recebimento") or ""),
                "status":           str(row.get("status") or ""),
            })
        return result
    except Exception:
        return []

# ── Auth ──────────────────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("login_page"))
        return f(*args, **kwargs)
    return wrapper

@app.route("/login", methods=["GET", "POST"])
def login_page():
    error = None
    if request.method == "POST":
        u = request.form.get("username", "").strip()
        p = request.form.get("password", "")
        if u in USERS and USERS[u]["password"] == p:
            session["user"] = u
            session["role"] = USERS[u]["role"]
            return redirect("/")
        error = "Usuário ou senha inválidos."
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# ── Pages ─────────────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def index():
    return render_template("index.html", user=session["user"])

@app.route("/folhas")
@login_required
def folhas_page():
    return render_template("folhas.html", user=session["user"])

# ── API ───────────────────────────────────────────────────────────────────────

@app.route("/api/medicoes")
@login_required
def api_list():
    with get_db() as conn:
        rows = conn.execute(
            "SELECT * FROM medicoes WHERE delete_requested=0 ORDER BY comp DESC,contrato_num,obra"
        ).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/medicoes", methods=["POST"])
@login_required
def api_create():
    b = request.json
    now = datetime.now().isoformat()
    id_ = str(uuid.uuid4())
    with get_db() as conn:
        conn.execute("""
            INSERT INTO medicoes(id,empresa,gestor,contrato_num,contrato_nome,
                obra,cod,comp,provisao,medicao,pedido,nf,venc_nf,
                retencao,impostos,status,obs,created_at,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (id_, b.get("empresa"), b.get("gestor"),
              b.get("contrato_num"), b.get("contrato_nome"),
              b.get("obra"), b.get("cod"), b.get("comp"),
              b.get("provisao", 0), b.get("medicao") or None,
              b.get("pedido"), b.get("nf"), b.get("venc_nf") or None,
              b.get("retencao") or None, b.get("impostos") or None,
              b.get("status", "previsto"), b.get("obs"), now, now))
    return jsonify({"id": id_}), 201

@app.route("/api/medicoes/<id>", methods=["PUT"])
@login_required
def api_update(id):
    b = request.json
    now = datetime.now().isoformat()
    with get_db() as conn:
        conn.execute("""
            UPDATE medicoes SET empresa=?,gestor=?,contrato_num=?,contrato_nome=?,
                obra=?,cod=?,comp=?,provisao=?,medicao=?,pedido=?,nf=?,venc_nf=?,
                retencao=?,impostos=?,status=?,obs=?,updated_at=?
            WHERE id=?
        """, (b.get("empresa"), b.get("gestor"),
              b.get("contrato_num"), b.get("contrato_nome"),
              b.get("obra"), b.get("cod"), b.get("comp"),
              b.get("provisao", 0), b.get("medicao") or None,
              b.get("pedido"), b.get("nf"), b.get("venc_nf") or None,
              b.get("retencao") or None, b.get("impostos") or None,
              b.get("status", "previsto"), b.get("obs"), now, id))
    return jsonify({"ok": True})

@app.route("/api/medicoes/<id>/solicitar-exclusao", methods=["POST"])
@login_required
def api_delete_request(id):
    now = datetime.now().isoformat()
    with get_db() as conn:
        row = conn.execute("SELECT * FROM medicoes WHERE id=?", (id,)).fetchone()
        if not row:
            return jsonify({"erro": "Não encontrado"}), 404
        row = dict(row)
        if session.get("role") == "admin":
            conn.execute("DELETE FROM medicoes WHERE id=?", (id,))
            return jsonify({"ok": True})
        conn.execute(
            "UPDATE medicoes SET delete_requested=1,delete_requested_by=?,delete_requested_at=? WHERE id=?",
            (session["user"], now, id))
        conn.execute("""
            INSERT INTO delete_requests(id,medicao_id,requested_by,requested_at,obra,contrato_num,contrato_nome,comp)
            VALUES(?,?,?,?,?,?,?,?)
        """, (str(uuid.uuid4()), id, session["user"], now,
              row.get("obra"), row.get("contrato_num"), row.get("contrato_nome"), row.get("comp")))
    return jsonify({"ok": True})

@app.route("/api/contratos")
@login_required
def api_contratos():
    with get_db() as conn:
        rows = conn.execute("SELECT * FROM contratos ORDER BY num").fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/contratos/<num>", methods=["PUT"])
@login_required
def api_update_contrato(num):
    b = request.json
    with get_db() as conn:
        conn.execute("INSERT OR REPLACE INTO contratos(num,nome,saldo) VALUES(?,?,?)",
                     (num, b.get("nome", ""), b.get("saldo", 0)))
    return jsonify({"ok": True})

@app.route("/api/folhas")
@login_required
def api_folhas():
    return jsonify(_read_controle())

@app.route("/api/delete-requests")
@login_required
def api_delete_requests():
    if session.get("role") != "admin":
        return jsonify({"erro": "Sem permissão"}), 403
    with get_db() as conn:
        rows = conn.execute("SELECT * FROM delete_requests ORDER BY requested_at DESC").fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/delete-requests/<id>/approve", methods=["POST"])
@login_required
def api_approve(id):
    if session.get("role") != "admin":
        return jsonify({"erro": "Sem permissão"}), 403
    with get_db() as conn:
        req = conn.execute("SELECT * FROM delete_requests WHERE id=?", (id,)).fetchone()
        if not req:
            return jsonify({"erro": "Não encontrado"}), 404
        conn.execute("DELETE FROM medicoes WHERE id=?", (req["medicao_id"],))
        conn.execute("DELETE FROM delete_requests WHERE id=?", (id,))
    return jsonify({"ok": True})

@app.route("/api/delete-requests/<id>/reject", methods=["POST"])
@login_required
def api_reject(id):
    if session.get("role") != "admin":
        return jsonify({"erro": "Sem permissão"}), 403
    with get_db() as conn:
        req = conn.execute("SELECT * FROM delete_requests WHERE id=?", (id,)).fetchone()
        if req:
            conn.execute("UPDATE medicoes SET delete_requested=0 WHERE id=?", (req["medicao_id"],))
        conn.execute("DELETE FROM delete_requests WHERE id=?", (id,))
    return jsonify({"ok": True})

@app.route("/api/importar/preview", methods=["POST"])
@login_required
def api_import_preview():
    file = request.files.get("arquivo")
    if not file:
        return jsonify({"erro": "Nenhum arquivo enviado"}), 400
    try:
        buf = BytesIO(file.read())
        df = pd.read_excel(buf, sheet_name="index", header=1, usecols="B:S")
        df.columns = [
            "empresa","contrato_num","contrato_nome","gestor",
            "obra","cod","comp","municipio",
            "provisao","valor_bruto_fat","pedido","nf",
            "protocolo","valor_bruto","retencao","impostos",
            "valor_liquido","data_recebimento",
        ]
        df = df[df["empresa"].notna()].copy()
        df["contrato_num"] = df["contrato_num"].astype(str).str.strip()
        df["comp"] = pd.to_datetime(df["comp"], errors="coerce").dt.strftime("%Y-%m")
        df["provisao"] = pd.to_numeric(df["provisao"], errors="coerce").fillna(0)
        df["_med"] = pd.to_numeric(df["valor_bruto_fat"], errors="coerce")

        def _st(row):
            nf  = str(row.get("nf","") or "").strip()
            ped = str(row.get("pedido","") or "").strip()
            med = row.get("_med")
            if nf  and nf  != "nan": return "nf_emitida"
            if ped and ped != "nan": return "aprovado"
            if pd.notna(med) and med > 0: return "medicao"
            return "previsto"

        rows = []
        for _, r in df.iterrows():
            ped = str(r.get("pedido","") or "").strip()
            nf  = str(r.get("nf","")    or "").strip()
            med = r.get("_med")
            rows.append({
                "empresa": r.get("empresa"), "gestor": r.get("gestor"),
                "contrato_num": r.get("contrato_num"), "contrato_nome": r.get("contrato_nome"),
                "obra": r.get("obra"), "cod": r.get("cod"), "comp": r.get("comp"),
                "provisao": float(r.get("provisao", 0)),
                "medicao": float(med) if pd.notna(med) else None,
                "pedido": ped if ped and ped != "nan" else None,
                "nf":     nf  if nf  and nf  != "nan" else None,
                "status": _st(r),
            })
        return jsonify({"total": len(rows), "preview": rows[:10], "rows": rows})
    except Exception as e:
        return jsonify({"erro": str(e)}), 400

@app.route("/api/importar/salvar", methods=["POST"])
@login_required
def api_import_save():
    rows = request.json.get("rows", [])
    now = datetime.now().isoformat()
    inseridas = 0
    with get_db() as conn:
        for r in rows:
            conn.execute("""
                INSERT INTO medicoes(id,empresa,gestor,contrato_num,contrato_nome,
                    obra,cod,comp,provisao,medicao,pedido,nf,status,created_at,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (str(uuid.uuid4()), r.get("empresa"), r.get("gestor"),
                  r.get("contrato_num"), r.get("contrato_nome"),
                  r.get("obra"), r.get("cod"), r.get("comp"),
                  r.get("provisao", 0), r.get("medicao"),
                  r.get("pedido"), r.get("nf"),
                  r.get("status","previsto"), now, now))
            inseridas += 1
    return jsonify({"inseridas": inseridas, "atualizadas": 0})

@app.route("/api/exportar")
@login_required
def api_exportar():
    with get_db() as conn:
        rows = conn.execute("SELECT * FROM medicoes WHERE delete_requested=0").fetchall()
    df = pd.DataFrame([dict(r) for r in rows])
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return send_file(buf, download_name="faturamento_export.xlsx",
                     as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(debug=True, host="0.0.0.0", port=port)
