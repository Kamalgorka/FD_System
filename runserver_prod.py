from waitress import serve
from core.wsgi import application

print("🚀 FD System Production Server Starting on http://0.0.0.0:8000")
serve(application, host="0.0.0.0", port=8000)