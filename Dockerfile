# Same image works locally (docker compose) and on Render (Render can deploy
# straight from this Dockerfile instead of the buildCommand/startCommand in
# render.yaml, if you ever switch it to "Docker" runtime there).
FROM python:3.13-slim

WORKDIR /app

# Install deps first so this layer stays cached while only app code changes.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY dataverse.py main.py ./

# Render sets $PORT itself; default to 8942 for local `docker run`.
ENV PORT=8942
EXPOSE 8942

CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT}"]
