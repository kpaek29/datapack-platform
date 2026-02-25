import os
from pathlib import Path

# Paths
BASE_DIR = Path(__file__).resolve().parent.parent

# Use persistent disk on Render, local dirs otherwise
if os.getenv("RENDER"):
    DATA_DIR = Path("/data")
else:
    DATA_DIR = BASE_DIR

UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "outputs"
TEMPLATE_DIR = BASE_DIR / "templates"  # Templates stay with code
USERS_FILE_PATH = DATA_DIR / "users.json"

# Create directories if they don't exist
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
TEMPLATE_DIR.mkdir(exist_ok=True)

# Security
SECRET_KEY = os.getenv("SECRET_KEY", "change-this-in-production-use-random-string")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 480  # 8 hours

# OpenAI (for AI-assisted analysis)
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
