# DataPack Platform

Secure data pack generation for Private Equity teams.

**Features:**
- 🔒 Secure login (JWT authentication)
- 📁 Upload Excel/CSV data files
- 📊 Automatic analysis (financial, customer data)
- 📑 Generate PPT presentations
- 📗 Generate Excel backups
- 🤝 Team collaboration with invite codes

## Quick Start

### 1. Install Dependencies

```bash
cd backend
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 2. Configure Environment

Create a `.env` file:
```bash
SECRET_KEY=your-random-secret-key-here
OPENAI_API_KEY=sk-...  # Optional, for AI features
```

### 3. Run Server

```bash
uvicorn backend.main:app --host 0.0.0.0 --port 8000
```

Access at `http://localhost:8000`

### 4. Default Login

- Username: `admin`
- Password: `changeme123`

**⚠️ Change this immediately in production!**

## Invite Codes

New users need an invite code to register. Default: `DATAPACK2024`

Change this in `backend/main.py` for production.

## Security Notes

- All uploads are stored in `uploads/` (gitignored)
- All outputs are stored in `outputs/` (gitignored)
- User passwords are hashed with bcrypt
- Sessions use JWT tokens (8-hour expiry)
- For production: use HTTPS, change default credentials, restrict CORS

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/auth/login` | POST | Login |
| `/api/auth/register` | POST | Register (needs invite code) |
| `/api/upload` | POST | Upload files |
| `/api/process/{session_id}` | POST | Process uploaded files |
| `/api/generate/{session_id}` | POST | Generate PPT/Excel |
| `/api/download/{session_id}/{filename}` | GET | Download output |
| `/api/sessions` | GET | List your sessions |

## Adding Analysis Types

Edit `backend/processor.py` to add new analysis patterns:
- Add keywords to detect data types
- Add analysis functions for each type

Edit `backend/generators.py` to customize outputs:
- Modify PPT slide layouts
- Customize Excel formatting
