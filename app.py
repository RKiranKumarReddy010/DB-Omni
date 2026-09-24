import os
import re
import sqlite3
import pandas as pd
from flask import Flask, request, jsonify

app = Flask(__name__)

DATABASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Database")
os.makedirs(DATABASE_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {"xlsx", "xls", "xlsm", "xlsb", "ods"}


# ─── Helpers ───────────────────────────────────────────────────────────────────

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


def get_db_path(db_name: str) -> str:
    """Resolve full path; db_name may or may not carry the .db extension."""
    if not db_name.endswith(".db"):
        db_name += ".db"
    return os.path.join(DATABASE_DIR, db_name)


def _safe_isna(v) -> bool:
    """Return True if v is NA/NaN/NaT, safely handles array-like values."""
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def excel_to_sqlite(file_path: str, db_name: str) -> dict:
    """
    Parse every sheet of an Excel file and write each sheet as a table
    inside Database/<db_name>.db.  Returns a summary dict.

    The ExcelFile handle is explicitly closed before this function returns
    so Windows releases the file lock and the caller can safely delete the
    temporary upload file.
    """
    db_path = get_db_path(db_name)
    xl = pd.ExcelFile(file_path, engine=pick_engine(file_path))
    summary = {"db_name": db_name, "db_file": f"{db_name}.db", "sheets": []}

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
                        # datetime / Timestamp / date
                        values.append(v.isoformat())
                    else:
                        # Convert numpy scalars to native Python types
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
        # IMPORTANT: close the ExcelFile handle so Windows releases the lock
        # on the temporary file before the caller tries to delete it.
        xl.close()
        conn.close()

    return summary


# ═══════════════════════════════════════════════════════════════════════════════
#  API ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

# ─── Upload ────────────────────────────────────────────────────────────────────
# POST /api/upload
#   Body: multipart/form-data  field "file" = Excel workbook
#   Creates Database/<filename>.db with one table per sheet.

@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file part in request"}), 400
    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400
    if not allowed_file(file.filename):
        return jsonify({"error": "Unsupported type. Accepted: .xlsx .xls .xlsm .xlsb .ods"}), 400

    db_name = sanitize_name(os.path.splitext(file.filename)[0])
    tmp_path = os.path.join(DATABASE_DIR, "__tmp__" + os.path.splitext(file.filename)[1])
    file.save(tmp_path)

    try:
        summary = excel_to_sqlite(tmp_path, db_name)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    return jsonify({"success": True, "summary": summary}), 201


# ─── Databases ─────────────────────────────────────────────────────────────────
# GET  /api/databases          – list all .db files
# DELETE /api/databases/<name> – delete a database file

@app.route("/api/databases", methods=["GET"])
def list_databases():
    dbs = []
    for fname in sorted(os.listdir(DATABASE_DIR)):
        if not fname.endswith(".db"):
            continue
        fpath = os.path.join(DATABASE_DIR, fname)
        conn = sqlite3.connect(fpath)
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
        tables = [r[0] for r in cur.fetchall()]
        conn.close()
        dbs.append({
            "name": fname,
            "size_bytes": os.path.getsize(fpath),
            "tables": tables,
        })
    return jsonify(dbs)


@app.route("/api/databases/<db_name>", methods=["DELETE"])
def delete_database(db_name):
    db_path = get_db_path(db_name)
    if not os.path.exists(db_path):
        return jsonify({"error": "Database not found"}), 404
    os.remove(db_path)
    return jsonify({"success": True, "deleted": db_name})


# ─── Tables ────────────────────────────────────────────────────────────────────
# GET /api/databases/<db>/tables                     – list tables
# GET /api/databases/<db>/tables/<table>/schema      – column definitions

@app.route("/api/databases/<db_name>/tables", methods=["GET"])
def list_tables(db_name):
    db_path = get_db_path(db_name)
    if not os.path.exists(db_path):
        return jsonify({"error": "Database not found"}), 404
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    tables = [r[0] for r in cur.fetchall()]
    conn.close()
    return jsonify({"tables": tables})


@app.route("/api/databases/<db_name>/tables/<table_name>/schema", methods=["GET"])
def get_schema(db_name, table_name):
    db_path = get_db_path(db_name)
    if not os.path.exists(db_path):
        return jsonify({"error": "Database not found"}), 404
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(f'PRAGMA table_info("{table_name}")')
    columns = [{"cid": r["cid"], "name": r["name"], "type": r["type"],
                "notnull": bool(r["notnull"]), "pk": bool(r["pk"])} for r in cur.fetchall()]
    conn.close()
    return jsonify({"table": table_name, "columns": columns})


# ─── CRUD : Read ───────────────────────────────────────────────────────────────
# GET /api/databases/<db>/tables/<table>/rows
#   Query params: page (default 1), limit (default 50, max 500), search (text)

@app.route("/api/databases/<db_name>/tables/<table_name>/rows", methods=["GET"])
def read_rows(db_name, table_name):
    db_path = get_db_path(db_name)
    if not os.path.exists(db_path):
        return jsonify({"error": "Database not found"}), 404

    page   = max(1, int(request.args.get("page", 1)))
    limit  = min(500, max(1, int(request.args.get("limit", 50))))
    search = request.args.get("search", "").strip()
    offset = (page - 1) * limit

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute(f'PRAGMA table_info("{table_name}")')
    columns = [{"name": r["name"], "type": r["type"]} for r in cur.fetchall()]

    where_clause, params_w = "", []
    if search:
        text_cols = [c["name"] for c in columns if "TEXT" in c["type"].upper() or c["type"] == ""]
        if text_cols:
            conds = " OR ".join([f'CAST("{c}" AS TEXT) LIKE ?' for c in text_cols])
            where_clause = f"WHERE {conds}"
            params_w = [f"%{search}%"] * len(text_cols)

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
        "rows": rows,
        "total": total,
        "page": page,
        "limit": limit,
        "pages": max(1, -(-total // limit)),
    })


# ─── CRUD : Create ─────────────────────────────────────────────────────────────
# POST /api/databases/<db>/tables/<table>/rows
#   Body: JSON object with column:value pairs  (omit "id")

@app.route("/api/databases/<db_name>/tables/<table_name>/rows", methods=["POST"])
def create_row(db_name, table_name):
    db_path = get_db_path(db_name)
    if not os.path.exists(db_path):
        return jsonify({"error": "Database not found"}), 404

    data = request.get_json()
    if not data:
        return jsonify({"error": "Request body must be JSON"}), 400

    data.pop("id", None)
    cols   = list(data.keys())
    values = list(data.values())
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


# ─── CRUD : Update ─────────────────────────────────────────────────────────────
# PUT /api/databases/<db>/tables/<table>/rows/<id>
#   Body: JSON object with the fields to update  (omit "id")

@app.route("/api/databases/<db_name>/tables/<table_name>/rows/<int:row_id>", methods=["PUT"])
def update_row(db_name, table_name, row_id):
    db_path = get_db_path(db_name)
    if not os.path.exists(db_path):
        return jsonify({"error": "Database not found"}), 404

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


# ─── CRUD : Delete ─────────────────────────────────────────────────────────────
# DELETE /api/databases/<db>/tables/<table>/rows/<id>

@app.route("/api/databases/<db_name>/tables/<table_name>/rows/<int:row_id>", methods=["DELETE"])
def delete_row(db_name, table_name, row_id):
    db_path = get_db_path(db_name)
    if not os.path.exists(db_path):
        return jsonify({"error": "Database not found"}), 404

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


# ─── SQL Console ───────────────────────────────────────────────────────────────
# POST /api/databases/<db>/query
#   Body: { "sql": "SELECT ..." }

@app.route("/api/databases/<db_name>/query", methods=["POST"])
def run_query(db_name):
    db_path = get_db_path(db_name)
    if not os.path.exists(db_path):
        return jsonify({"error": "Database not found"}), 404

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
    print("DB-Omni API  ->  http://127.0.0.1:5000")
    app.run(debug=True, port=5000)
