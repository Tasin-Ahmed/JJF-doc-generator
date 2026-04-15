# Procfile — Tells Railway how to start the application.
# "web:" is the process type Railway looks for to expose a public URL.
# uvicorn runs the FastAPI app defined in interfaces/telegram_bot.py as fastapi_app.
# --host 0.0.0.0 binds to all interfaces (required by Railway).
# --port $PORT uses Railway's dynamically assigned port.
web: uvicorn interfaces.telegram_bot:fastapi_app --host 0.0.0.0 --port $PORT
