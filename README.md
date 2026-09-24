# DB-Omni

A lightweight **REST API** built with Flask that converts Excel workbooks into
SQLite databases and exposes full **CRUD** operations on every table.

---

## Project Structure

```
DB-Omni/
├── app.py            # Flask API application
├── requirements.txt  # Python dependencies
├── Database/         # Auto-created; stores all .db files
└── README.md
```

---

## Setup

### 1. Clone / navigate to the project

```bash
cd DB-Omni
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the server

```bash
python app.py
```

The API will be available at **`http://127.0.0.1:5000`**.

---

## How it works

1. Upload any Excel file (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.ods`).
2. Each **sheet** becomes a **SQLite table** inside `Database/<filename>.db`.
3. Column types are inferred automatically (`INTEGER`, `REAL`, `TEXT`).
4. Every table gets an auto-increment `id` primary key.
5. Use the CRUD endpoints to manage rows, or run raw SQL via the query endpoint.

---

## API Reference

### Upload

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/api/upload` | Upload an Excel file and create a database |

**Request** — `multipart/form-data`

| Field | Type | Description |
|-------|------|-------------|
| `file` | File | Excel workbook (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.ods`) |

**Response `201`**
```json
{
  "success": true,
  "summary": {
    "db_name": "sales_data",
    "db_file": "sales_data.db",
    "sheets": [
      {
        "sheet": "Q1",
        "table": "Q1",
        "rows": 120,
        "columns": ["product", "revenue", "units"]
      }
    ]
  }
}
```

---

### Databases

| Method | Endpoint | Description |
|--------|----------|-------------|
| `GET` | `/api/databases` | List all databases |
| `DELETE` | `/api/databases/<db_name>` | Delete a database file |

**GET `/api/databases` — Response**
```json
[
  {
    "name": "sales_data.db",
    "size_bytes": 32768,
    "tables": ["Q1", "Q2", "Q3"]
  }
]
```

> **Note:** `<db_name>` can be provided with or without the `.db` extension.

---

### Tables

| Method | Endpoint | Description |
|--------|----------|-------------|
| `GET` | `/api/databases/<db_name>/tables` | List all tables in a database |
| `GET` | `/api/databases/<db_name>/tables/<table_name>/schema` | Get column definitions |

**GET schema — Response**
```json
{
  "table": "Q1",
  "columns": [
    { "cid": 0, "name": "id",      "type": "INTEGER", "notnull": false, "pk": true  },
    { "cid": 1, "name": "product", "type": "TEXT",    "notnull": false, "pk": false },
    { "cid": 2, "name": "revenue", "type": "REAL",    "notnull": false, "pk": false }
  ]
}
```

---

### CRUD — Rows

#### Read (paginated + searchable)

```
GET /api/databases/<db_name>/tables/<table_name>/rows
```

| Query Param | Default | Description |
|-------------|---------|-------------|
| `page` | `1` | Page number |
| `limit` | `50` | Rows per page (max 500) |
| `search` | — | Full-text search across TEXT columns |

**Response**
```json
{
  "columns": [{"name": "id", "type": "INTEGER"}, {"name": "product", "type": "TEXT"}],
  "rows": [{"id": 1, "product": "Widget A", "revenue": 4500.0}],
  "total": 120,
  "page": 1,
  "limit": 50,
  "pages": 3
}
```

---

#### Create

```
POST /api/databases/<db_name>/tables/<table_name>/rows
```

**Request body** (JSON — omit `id`):
```json
{
  "product": "Widget B",
  "revenue": 6200.0,
  "units": 80
}
```

**Response `201`**
```json
{
  "success": true,
  "row": { "id": 121, "product": "Widget B", "revenue": 6200.0, "units": 80 }
}
```

---

#### Update

```
PUT /api/databases/<db_name>/tables/<table_name>/rows/<row_id>
```

**Request body** (JSON — only fields to change):
```json
{
  "revenue": 7500.0
}
```

**Response**
```json
{
  "success": true,
  "row": { "id": 121, "product": "Widget B", "revenue": 7500.0, "units": 80 }
}
```

---

#### Delete

```
DELETE /api/databases/<db_name>/tables/<table_name>/rows/<row_id>
```

**Response**
```json
{
  "success": true,
  "deleted_id": 121
}
```

---

### SQL Console

```
POST /api/databases/<db_name>/query
```

**Request body**
```json
{
  "sql": "SELECT product, SUM(revenue) AS total FROM Q1 GROUP BY product"
}
```

**Response — SELECT**
```json
{
  "columns": ["product", "total"],
  "rows": [{"product": "Widget A", "total": 45000.0}],
  "count": 1
}
```

**Response — DML (INSERT / UPDATE / DELETE)**
```json
{
  "message": "3 row(s) affected"
}
```

---

## Error Responses

All errors return a JSON body with an `"error"` key and an appropriate HTTP status code.

| Status | Meaning |
|--------|---------|
| `400` | Bad request (missing fields, invalid SQL, etc.) |
| `404` | Database or row not found |
| `500` | Internal server error (file parsing failure, etc.) |

```json
{ "error": "Row 99 not found" }
```

---

## Quick cURL Examples

```bash
# Upload Excel
curl -X POST http://127.0.0.1:5000/api/upload \
     -F "file=@sales_data.xlsx"

# List databases
curl http://127.0.0.1:5000/api/databases

# Read rows (page 2, 25 per page)
curl "http://127.0.0.1:5000/api/databases/sales_data/tables/Q1/rows?page=2&limit=25"

# Search rows
curl "http://127.0.0.1:5000/api/databases/sales_data/tables/Q1/rows?search=Widget"

# Create a row
curl -X POST http://127.0.0.1:5000/api/databases/sales_data/tables/Q1/rows \
     -H "Content-Type: application/json" \
     -d '{"product":"Widget C","revenue":3000,"units":40}'

# Update a row
curl -X PUT http://127.0.0.1:5000/api/databases/sales_data/tables/Q1/rows/1 \
     -H "Content-Type: application/json" \
     -d '{"revenue":9999}'

# Delete a row
curl -X DELETE http://127.0.0.1:5000/api/databases/sales_data/tables/Q1/rows/1

# Run custom SQL
curl -X POST http://127.0.0.1:5000/api/databases/sales_data/query \
     -H "Content-Type: application/json" \
     -d '{"sql":"SELECT * FROM Q1 LIMIT 5"}'

# Delete a database
curl -X DELETE http://127.0.0.1:5000/api/databases/sales_data
```

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `flask` | Web framework / API server |
| `pandas` | Excel parsing & data manipulation |
| `openpyxl` | Read `.xlsx`, `.xlsm`, `.xlsb` files |
| `xlrd` | Read legacy `.xls` files |
