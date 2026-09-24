"""
DB-Omni API
===========
Authentication flow:
  1. POST /api/auth/login  { "user_id": "...", "password": "..." }
     → queries Firestore `playground_users` collection (same as OmniTensors_dashboard)
     → returns a signed JWT carrying { user_id, privileges }

  2. Every other endpoint requires:
        Authorization: Bearer <token>
     The JWT is decoded to extract `user_id`, which is used as the SQLite
     database filename:  Database/<user_id>.db

Database scoping:
  Each authenticated user gets their own isolated database file.
  No user can access another user's data.
"""

import os
import re
import sqlite3
import datetime
import functools

import jwt
import pandas as pd
import firebase_admin
from firebase_admin import credentials, firestore
from flask import Flask, request, jsonify
from dotenv import load_dotenv

# ─── Bootstrap ─────────────────────────────────────────────────────────────────

load_dotenv()

app = Flask(__name__)

DATABASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Database")
os.makedirs(DATABASE_DIR, exist_ok=True)

# Firebase is initialised lazily inside get_firestore().
# Nothing runs at import time, so the server always starts successfully
# even if serviceAccountKey.json is not yet present.
_firebase_app = None
_fs           = None


def get_firestore():
    """
    Lazily initialise Firebase Admin SDK and return a Firestore client.

    Credential resolution order:
      1. FIREBASE_CREDENTIALS env var  — full JSON string (use on Vercel)
      2. Any *.json file in Private/   — local dev convenience
      3. serviceAccountKey.json        — local dev fallback

    Raises RuntimeError with setup instructions if nothing is found.
    """
    global _firebase_app, _fs

    if _fs is not None:
        return _fs

    root = os.path.dirname(os.path.abspath(__file__))
    cred = None

    # ── 1. Env var: full JSON string (Vercel / CI) ─────────────────────────
    creds_json = os.getenv("FIREBASE_CREDENTIALS", "").strip()
    if creds_json:
        try:
            cred_dict = json.loads(creds_json)
            cred = credentials.Certificate(cred_dict)
        except Exception as e:
            raise RuntimeError(f"FIREBASE_CREDENTIALS env var is not valid JSON: {e}")

    # ── 2 & 3. File-based (local dev) ──────────────────────────────────────
    if cred is None:
        candidates = []

        env_path = os.getenv("FIREBASE_SERVICE_ACCOUNT_JSON", "")
        if env_path:
            candidates.append(os.path.join(root, env_path))

        private_dir = os.path.join(root, "Private")
        if os.path.isdir(private_dir):
            for fname in sorted(os.listdir(private_dir)):
                if fname.endswith(".json"):
                    candidates.append(os.path.join(private_dir, fname))

        candidates.append(os.path.join(root, "serviceAccountKey.json"))

        sa_path = next((p for p in candidates if os.path.exists(p)), None)

        if sa_path is None:
            raise RuntimeError(
                "Firebase credentials not found. "
                "On Vercel: set the FIREBASE_CREDENTIALS env var to the full JSON content "
                "of your service account key. "
                "Locally: place the JSON file in the Private/ folder."
            )
        cred = credentials.Certificate(sa_path)

    if not firebase_admin._apps:
        _firebase_app = firebase_admin.initialize_app(cred)

    _fs = firestore.client()
    return _fs

# ─── JWT config ────────────────────────────────────────────────────────────────

JWT_SECRET    = os.getenv("JWT_SECRET", "change-me-to-a-long-random-secret")
JWT_ALGORITHM = os.getenv("JWT_ALGORITHM", "HS256")
JWT_EXPIRE    = int(os.getenv("JWT_EXPIRE_SECONDS", 28800))  # 8 hours

ALLOWED_EXTENSIONS = {"xlsx", "xls", "xlsm", "xlsb", "ods"}


# ═══════════════════════════════════════════════════════════════════════════════
#  Helpers
# ═══════════════════════════════════════════════════════════════════════════════

def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def sanitize_name(name: str) -> str:
    """Return a safe SQLite identifier from an arbitrary string."""
    name = str(name).strip()
    name = re.sub(r"[^\w]", "_", name)
    if name and name[0].isdigit():
        name = "_" + name
    return name or "unnamed"


def infer_dtype(series: pd.Series) -> str:
    """Map a pandas Series dtype to a SQLite column type."""
    if pd.api.types.is_integer_dtype(series):
        return "INTEGER"
    if pd.api.types.is_float_dtype(series):
        return "REAL"
    return "TEXT"


def pick_engine(path: str) -> str:
    return "xlrd" if path.rsplit(".", 1)[-1].lower() == "xls" else "openpyxl"


def _safe_isna(v) -> bool:
    """Return True if v is NA/NaN/NaT; handles scalars and array-likes safely."""
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def get_user_db_path(user_id: str) -> str:
    """
    Each user gets their own database file: Database/<user_id>.db
    user_id is sanitised so it is always a safe filename.
    """
    safe_id = re.sub(r"[^\w\-]", "_", user_id)
    return os.path.join(DATABASE_DIR, f"{safe_id}.db")


def excel_to_sqlite(file_path: str, db_path: str) -> dict:
    """
    Parse every sheet of an Excel file and write each sheet as a table
    inside the given db_path.  Returns a summary dict.
    """
    xl = pd.ExcelFile(file_path, engine=pick_engine(file_path))
    summary = {"db_file": os.path.basename(db_path), "sheets": []}

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    try:
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            df.dropna(how="all", inplace=True)
            df.dropna(axis=1, how="all", inplace=True)

            if df.empty:
                continue

            # Sanitize & deduplicate column names
            cols, seen = [], {}
            for c in df.columns:
                c = sanitize_name(c)
                if c in seen:
                    seen[c] += 1
                    c = f"{c}_{seen[c]}"
                else:
                    seen[c] = 0
                cols.append(c)
            df.columns = cols

            table_name = sanitize_name(sheet)
            col_defs = [f'"{c}" {infer_dtype(df[c])}' for c in df.columns]

            cur.execute(
                f'CREATE TABLE IF NOT EXISTS "{table_name}" '
                f'(id INTEGER PRIMARY KEY AUTOINCREMENT, {", ".join(col_defs)})'
            )

            rows_inserted = 0
            for _, row in df.iterrows():
                values = []
                for v in row:
                    if _safe_isna(v):
                        values.append(None)
                    elif hasattr(v, "isoformat"):
                        values.append(v.isoformat())
                    else:
                        values.append(v.item() if hasattr(v, "item") else v)
                ph = ", ".join(["?"] * len(df.columns))
                col_list = ", ".join([f'"{c}"' for c in df.columns])
                cur.execute(
                    f'INSERT INTO "{table_name}" ({col_list}) VALUES ({ph})', values
                )
                rows_inserted += 1

            conn.commit()
            summary["sheets"].append({
                "sheet": sheet,
                "table": table_name,
                "rows": rows_inserted,
                "columns": list(df.columns),
            })
    finally:
        # Release file handle so Windows can delete the temp file
        xl.close()
        conn.close()

    return summary


# ═══════════════════════════════════════════════════════════════════════════════
#  JWT helpers
# ═══════════════════════════════════════════════════════════════════════════════

def create_token(user_id: str, privileges: str) -> str:
    payload = {
        "user_id":    user_id,
        "privileges": privileges,
        "exp": datetime.datetime.utcnow() + datetime.timedelta(seconds=JWT_EXPIRE),
        "iat": datetime.datetime.utcnow(),
    }
    return jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALGORITHM)


def decode_token(token: str) -> dict:
    """Decode and validate JWT. Raises jwt.exceptions.* on failure."""
    return jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])


def require_auth(f):
    """
    Decorator: extracts Bearer token from Authorization header,
    decodes it, and injects `current_user` dict into the view function.
    """
    @functools.wraps(f)
    def wrapper(*args, **kwargs):
        auth_header = request.headers.get("Authorization", "")
        if not auth_header.startswith("Bearer "):
            return jsonify({"error": "Missing or malformed Authorization header"}), 401
        token = auth_header[7:]
        try:
            payload = decode_token(token)
        except jwt.ExpiredSignatureError:
            return jsonify({"error": "Token has expired. Please log in again."}), 401
        except jwt.InvalidTokenError as e:
            return jsonify({"error": f"Invalid token: {e}"}), 401

        return f(*args, current_user=payload, **kwargs)
    return wrapper


# ═══════════════════════════════════════════════════════════════════════════════
#  AUTH ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

# POST /api/auth/login
# Body: { "user_id": "kiran.kumar", "password": "••••••••" }
# Queries the same Firestore `playground_users` collection used by OmniTensors.
# Returns: { "token": "<jwt>", "user_id": "...", "privileges": "..." }

@app.route("/api/auth/login", methods=["POST"])
def login():
    body = request.get_json() or {}
    user_id  = body.get("user_id", "").strip()
    password = body.get("password", "").strip()

    if not user_id or not password:
        return jsonify({"error": "user_id and password are required"}), 400

    try:
        fs = get_firestore()
        docs = (
            fs.collection("playground_users")
            .where("userId", "==", user_id)
            .where("password", "==", password)
            .limit(1)
            .get()
        )
    except RuntimeError as exc:
        # serviceAccountKey.json is missing — return a setup guide
        return jsonify({"error": str(exc)}), 503
    except Exception as exc:
        return jsonify({"error": f"Firestore error: {exc}"}), 500

    if not docs:
        return jsonify({"error": "Invalid user_id or password"}), 401

    user_doc   = docs[0].to_dict()
    privileges = user_doc.get("privileges", "Interactive Analyst")
    token      = create_token(user_id, privileges)

    return jsonify({
        "token":      token,
        "user_id":    user_id,
        "privileges": privileges,
        "expires_in": JWT_EXPIRE,
    })


# GET /api/auth/me  → return current user info from token
@app.route("/api/auth/me", methods=["GET"])
@require_auth
def me(current_user):
    return jsonify({
        "user_id":    current_user["user_id"],
        "privileges": current_user["privileges"],
    })


# ═══════════════════════════════════════════════════════════════════════════════
#  UPLOAD
# ═══════════════════════════════════════════════════════════════════════════════

# POST /api/upload
# Header: Authorization: Bearer <token>
# Body: multipart/form-data  field "file" = Excel workbook
# Saves sheets as tables inside Database/<user_id>.db

@app.route("/api/upload", methods=["POST"])
@require_auth
def upload(current_user):
    if "file" not in request.files:
        return jsonify({"error": "No file part in request"}), 400
    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400
    if not allowed_file(file.filename):
        return jsonify({"error": "Unsupported type. Accepted: .xlsx .xls .xlsm .xlsb .ods"}), 400

    user_id = current_user["user_id"]
    db_path = get_user_db_path(user_id)
    ext     = os.path.splitext(file.filename)[1]
    tmp_path = os.path.join(DATABASE_DIR, f"__tmp_{user_id}__{ext}")
    file.save(tmp_path)

    try:
        summary = excel_to_sqlite(tmp_path, db_path)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    summary["user_id"] = user_id
    return jsonify({"success": True, "summary": summary}), 201


# ═══════════════════════════════════════════════════════════════════════════════
#  DATABASE INFO
# ═══════════════════════════════════════════════════════════════════════════════

# GET /api/database/info  → size, table list for the current user's DB

@app.route("/api/database/info", methods=["GET"])
@require_auth
def db_info(current_user):
    db_path = get_user_db_path(current_user["user_id"])
    if not os.path.exists(db_path):
        return jsonify({"error": "No database found. Upload an Excel file first."}), 404

    conn = sqlite3.connect(db_path)
    cur  = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    tables = [r[0] for r in cur.fetchall()]
    conn.close()

    return jsonify({
        "user_id":    current_user["user_id"],
        "db_file":    os.path.basename(db_path),
        "size_bytes": os.path.getsize(db_path),
        "tables":     tables,
    })


# DELETE /api/database  → wipe the current user's entire database file

@app.route("/api/database", methods=["DELETE"])
@require_auth
def delete_database(current_user):
    db_path = get_user_db_path(current_user["user_id"])
    if not os.path.exists(db_path):
        return jsonify({"error": "No database found"}), 404
    os.remove(db_path)
    return jsonify({"success": True, "message": "Your database has been deleted."})


# ═══════════════════════════════════════════════════════════════════════════════
#  TABLES
# ═══════════════════════════════════════════════════════════════════════════════

# GET /api/tables  → list tables

@app.route("/api/tables", methods=["GET"])
@require_auth
def list_tables(current_user):
    db_path = get_user_db_path(current_user["user_id"])
    if not os.path.exists(db_path):
        return jsonify({"error": "No database found. Upload an Excel file first."}), 404

    conn = sqlite3.connect(db_path)
    cur  = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    tables = [r[0] for r in cur.fetchall()]
    conn.close()
    return jsonify({"tables": tables})


# GET /api/tables/<table>/schema  → column definitions

@app.route("/api/tables/<table_name>/schema", methods=["GET"])
@require_auth
def get_schema(table_name, current_user):
    db_path = get_user_db_path(current_user["user_id"])
    if not os.path.exists(db_path):
        return jsonify({"error": "No database found"}), 404

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur  = conn.cursor()
    cur.execute(f'PRAGMA table_info("{table_name}")')
    columns = [
        {"cid": r["cid"], "name": r["name"], "type": r["type"],
         "notnull": bool(r["notnull"]), "pk": bool(r["pk"])}
        for r in cur.fetchall()
    ]
    conn.close()
    return jsonify({"table": table_name, "columns": columns})


# ═══════════════════════════════════════════════════════════════════════════════
#  CRUD – ROWS
# ═══════════════════════════════════════════════════════════════════════════════

# GET /api/tables/<table>/rows?page=1&limit=50&search=...  → paginated read

@app.route("/api/tables/<table_name>/rows", methods=["GET"])
@require_auth
def read_rows(table_name, current_user):
    db_path = get_user_db_path(current_user["user_id"])
    if not os.path.exists(db_path):
        return jsonify({"error": "No database found"}), 404

    page   = max(1, int(request.args.get("page", 1)))
    limit  = min(500, max(1, int(request.args.get("limit", 50))))
    search = request.args.get("search", "").strip()
    offset = (page - 1) * limit

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur  = conn.cursor()

    cur.execute(f'PRAGMA table_info("{table_name}")')
    columns = [{"name": r["name"], "type": r["type"]} for r in cur.fetchall()]

    where_clause, params_w = "", []
    if search:
        text_cols = [c["name"] for c in columns if "TEXT" in c["type"].upper() or c["type"] == ""]
        if text_cols:
            conds         = " OR ".join([f'CAST("{c}" AS TEXT) LIKE ?' for c in text_cols])
            where_clause  = f"WHERE {conds}"
            params_w      = [f"%{search}%"] * len(text_cols)

    cur.execute(f'SELECT COUNT(*) FROM "{table_name}" {where_clause}', params_w)
    total = cur.fetchone()[0]

    cur.execute(
        f'SELECT * FROM "{table_name}" {where_clause} LIMIT ? OFFSET ?',
        params_w + [limit, offset],
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()

    return jsonify({
        "columns": columns,
        "rows":    rows,
        "total":   total,
        "page":    page,
        "limit":   limit,
        "pages":   max(1, -(-total // limit)),
    })


# POST /api/tables/<table>/rows  → create row

@app.route("/api/tables/<table_name>/rows", methods=["POST"])
@require_auth
def create_row(table_name, current_user):
    db_path = get_user_db_path(current_user["user_id"])
    if not os.path.exists(db_path):
        return jsonify({"error": "No database found"}), 404

    data = request.get_json()
    if not data:
        return jsonify({"error": "Request body must be JSON"}), 400

    data.pop("id", None)
    cols    = list(data.keys())
    values  = list(data.values())
    col_sql = ", ".join([f'"{c}"' for c in cols])
    ph      = ", ".join(["?"] * len(cols))

    try:
        conn = sqlite3.connect(db_path)
        cur  = conn.cursor()
        cur.execute(f'INSERT INTO "{table_name}" ({col_sql}) VALUES ({ph})', values)
        conn.commit()
        new_id = cur.lastrowid
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute(f'SELECT * FROM "{table_name}" WHERE id = ?', (new_id,))
        row = dict(cur.fetchone())
        conn.close()
        return jsonify({"success": True, "row": row}), 201
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


# PUT /api/tables/<table>/rows/<id>  → update row

@app.route("/api/tables/<table_name>/rows/<int:row_id>", methods=["PUT"])
@require_auth
def update_row(table_name, row_id, current_user):
    db_path = get_user_db_path(current_user["user_id"])
    if not os.path.exists(db_path):
        return jsonify({"error": "No database found"}), 404

    data = request.get_json()
    if not data:
        return jsonify({"error": "Request body must be JSON"}), 400

    data.pop("id", None)
    set_clause = ", ".join([f'"{k}" = ?' for k in data.keys()])
    values     = list(data.values()) + [row_id]

    try:
        conn = sqlite3.connect(db_path)
        cur  = conn.cursor()
        cur.execute(f'UPDATE "{table_name}" SET {set_clause} WHERE id = ?', values)
        conn.commit()
        if cur.rowcount == 0:
            conn.close()
            return jsonify({"error": f"Row {row_id} not found"}), 404
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute(f'SELECT * FROM "{table_name}" WHERE id = ?', (row_id,))
        row = dict(cur.fetchone())
        conn.close()
        return jsonify({"success": True, "row": row})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


# DELETE /api/tables/<table>/rows/<id>  → delete row

@app.route("/api/tables/<table_name>/rows/<int:row_id>", methods=["DELETE"])
@require_auth
def delete_row(table_name, row_id, current_user):
    db_path = get_user_db_path(current_user["user_id"])
    if not os.path.exists(db_path):
        return jsonify({"error": "No database found"}), 404

    try:
        conn = sqlite3.connect(db_path)
        cur  = conn.cursor()
        cur.execute(f'DELETE FROM "{table_name}" WHERE id = ?', (row_id,))
        conn.commit()
        affected = cur.rowcount
        conn.close()
        if affected == 0:
            return jsonify({"error": f"Row {row_id} not found"}), 404
        return jsonify({"success": True, "deleted_id": row_id})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


# ═══════════════════════════════════════════════════════════════════════════════
#  SQL CONSOLE
# ═══════════════════════════════════════════════════════════════════════════════

# POST /api/query
# Body: { "sql": "SELECT ..." }

@app.route("/api/query", methods=["POST"])
@require_auth
def run_query(current_user):
    db_path = get_user_db_path(current_user["user_id"])
    if not os.path.exists(db_path):
        return jsonify({"error": "No database found"}), 404

    body = request.get_json() or {}
    sql  = body.get("sql", "").strip()
    if not sql:
        return jsonify({"error": "Field 'sql' is required"}), 400

    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cur  = conn.cursor()
        cur.execute(sql)
        if cur.description:
            columns = [d[0] for d in cur.description]
            rows    = [dict(r) for r in cur.fetchall()]
            conn.close()
            return jsonify({"columns": columns, "rows": rows, "count": len(rows)})
        conn.commit()
        affected = cur.rowcount
        conn.close()
        return jsonify({"message": f"{affected} row(s) affected"})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


# ───────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.getenv("FLASK_PORT", 5000))
    print(f"DB-Omni API -> http://127.0.0.1:{port}")
    app.run(debug=True, port=port)
