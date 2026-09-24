# DB-Omni

A **per-user, authenticated REST API** built with Flask that:
- Verifies users against the **Firebase `playground_users` Firestore collection** (same auth as OmniTensors_dashboard)
- Issues a **JWT token** on login
- Converts uploaded Excel workbooks into **SQLite databases scoped to the logged-in user** (`Database/<user_id>.db`)
- Exposes full **CRUD + SQL console** endpoints protected by Bearer token

---

## Project Structure

```
DB-Omni/
├── app.py                  # Flask API (auth + CRUD)
├── requirements.txt        # Python dependencies
├── .env                    # Configuration (JWT secret, Firebase project, etc.)
├── serviceAccountKey.json  # Firebase service-account key (you must add this)
├── Database/               # Auto-created; one .db file per user_id
└── README.md
```

---

## Setup

### 1. Firebase Service Account Key

Download your key from:
> **Firebase Console → Project Settings → Service Accounts → Generate new private key**

Save it as `serviceAccountKey.json` in the project root.

### 2. Configure `.env`

```env
FIREBASE_SERVICE_ACCOUNT_JSON=serviceAccountKey.json
JWT_SECRET=your-long-random-secret-here
JWT_EXPIRE_SECONDS=28800
FLASK_PORT=5000
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Run the server

```bash
python app.py
```

API available at **`http://127.0.0.1:5000`**

---

## Authentication Flow

```
Client                              DB-Omni API              Firestore
  │                                     │                       │
  │── POST /api/auth/login ────────────>│                       │
  │   { user_id, password }             │── query playground_users ──>│
  │                                     │<─── doc { userId, password, privileges } ──│
  │<── { token, user_id, privileges } ──│                       │
  │                                     │                       │
  │── POST /api/upload ─────────────────│  (Bearer token)       │
  │   Authorization: Bearer <token>     │                       │
  │                                     │  saves to Database/<user_id>.db
  │<── 201 { summary }  ────────────────│                       │
```

> Users are created by the **admin** in `OmniTensors_dashboard /admin` — the same `playground_users` Firestore collection is queried here.

---

## API Reference

### Auth

| Method | Endpoint | Auth | Description |
|--------|----------|------|-------------|
| `POST` | `/api/auth/login` | None | Login with playground credentials → JWT |
| `GET` | `/api/auth/me` | Bearer | Return current user info |

**POST `/api/auth/login`**
```json
// Request
{ "user_id": "kiran.kumar", "password": "••••••••" }

// Response 200
{
  "token": "<jwt>",
  "user_id": "kiran.kumar",
  "privileges": "Interactive Analyst",
  "expires_in": 28800
}
```

---

### Upload

| Method | Endpoint | Auth | Description |
|--------|----------|------|-------------|
| `POST` | `/api/upload` | Bearer | Upload Excel → create/append `Database/<user_id>.db` |

**Request** — `multipart/form-data`, field `file`

**Response `201`**
```json
{
  "success": true,
  "summary": {
    "user_id": "kiran.kumar",
    "db_file": "kiran_kumar.db",
    "sheets": [
      { "sheet": "Q1", "table": "Q1", "rows": 120, "columns": ["product","revenue"] }
    ]
  }
}
```

---

### Database Info

| Method | Endpoint | Auth | Description |
|--------|----------|------|-------------|
| `GET` | `/api/database/info` | Bearer | Get DB file info + table list |
| `DELETE` | `/api/database` | Bearer | Delete the user's entire database |

---

### Tables

| Method | Endpoint | Auth | Description |
|--------|----------|------|-------------|
| `GET` | `/api/tables` | Bearer | List tables |
| `GET` | `/api/tables/<table>/schema` | Bearer | Column definitions |

---

### CRUD — Rows

| Method | Endpoint | Auth | Description |
|--------|----------|------|-------------|
| `GET` | `/api/tables/<table>/rows` | Bearer | Read rows (paginated + search) |
| `POST` | `/api/tables/<table>/rows` | Bearer | Create row |
| `PUT` | `/api/tables/<table>/rows/<id>` | Bearer | Update row |
| `DELETE` | `/api/tables/<table>/rows/<id>` | Bearer | Delete row |

**Read query params**

| Param | Default | Description |
|-------|---------|-------------|
| `page` | `1` | Page number |
| `limit` | `50` | Rows per page (max 500) |
| `search` | — | Full-text search across TEXT columns |

---

### SQL Console

| Method | Endpoint | Auth | Description |
|--------|----------|------|-------------|
| `POST` | `/api/query` | Bearer | Run raw SQL against the user's DB |

**Request**
```json
{ "sql": "SELECT * FROM Q1 WHERE revenue > 5000" }
```

---

## Error Responses

| Status | Meaning |
|--------|---------|
| `400` | Bad request / invalid SQL |
| `401` | Missing token / expired token / wrong credentials |
| `404` | Database or row not found |
| `500` | Internal error (file parse, Firestore, etc.) |

```json
{ "error": "Token has expired. Please log in again." }
```

---

## Quick cURL Examples

```bash
# 1. Login
TOKEN=$(curl -s -X POST http://127.0.0.1:5000/api/auth/login \
  -H "Content-Type: application/json" \
  -d '{"user_id":"kiran.kumar","password":"yourpass"}' | python -c "import sys,json; print(json.load(sys.stdin)['token'])")

# 2. Upload Excel
curl -X POST http://127.0.0.1:5000/api/upload \
  -H "Authorization: Bearer $TOKEN" \
  -F "file=@sales_data.xlsx"

# 3. List tables
curl http://127.0.0.1:5000/api/tables \
  -H "Authorization: Bearer $TOKEN"

# 4. Read rows (page 1, 25 per page)
curl "http://127.0.0.1:5000/api/tables/Q1/rows?page=1&limit=25" \
  -H "Authorization: Bearer $TOKEN"

# 5. Search rows
curl "http://127.0.0.1:5000/api/tables/Q1/rows?search=Widget" \
  -H "Authorization: Bearer $TOKEN"

# 6. Create a row
curl -X POST http://127.0.0.1:5000/api/tables/Q1/rows \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"product":"Widget C","revenue":3000,"units":40}'

# 7. Update a row
curl -X PUT http://127.0.0.1:5000/api/tables/Q1/rows/1 \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"revenue":9999}'

# 8. Delete a row
curl -X DELETE http://127.0.0.1:5000/api/tables/Q1/rows/1 \
  -H "Authorization: Bearer $TOKEN"

# 9. Run custom SQL
curl -X POST http://127.0.0.1:5000/api/query \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"sql":"SELECT product, SUM(revenue) FROM Q1 GROUP BY product"}'

# 10. Delete your entire database
curl -X DELETE http://127.0.0.1:5000/api/database \
  -H "Authorization: Bearer $TOKEN"
```

---

## Dependencies

| Package | Purpose |
|---------|---------|
| `flask` | Web framework |
| `pandas` | Excel parsing |
| `openpyxl` | Read `.xlsx`, `.xlsm`, `.xlsb` |
| `xlrd` | Read legacy `.xls` |
| `firebase-admin` | Query Firestore `playground_users` |
| `PyJWT` | Issue & verify JWT tokens |
| `python-dotenv` | Load `.env` variables |
