# Google Maps Lead Scraper — Render API

This project converts the original command-line scraper into an HTTP API.

## Files

- `app.py` — FastAPI HTTP API
- `scraper.py` — Playwright + Excel scraping logic
- `requirements.txt` — Python dependencies
- `Dockerfile` — installs Chromium and runs the API
- `render.yaml` — Render deployment configuration

## Run locally

Install dependencies:

```bash
pip install -r requirements.txt
playwright install chromium
```

Start the API:

```bash
uvicorn app:app --reload
```

Open:

```text
http://127.0.0.1:8000/docs
```

FastAPI provides an interactive API page there.

## Scrape request

POST `/scrape`

Example JSON:

```json
{
  "query": "dentists",
  "location": "Vellore, Tamil Nadu",
  "max_results": 50,
  "scroll_pause": 2
}
```

The response is an Excel file named `google_maps_businesses.xlsx`.

## Deploy to Render

1. Put these files in a GitHub repository.
2. In Render, create a new Blueprint / Web Service from the repository.
3. Render will use `render.yaml`, or you can manually select Docker.
4. Deploy.
5. Your endpoint will be:

```text
https://YOUR-SERVICE-NAME.onrender.com/scrape
```

## Test with curl

```bash
curl -X POST "https://YOUR-SERVICE-NAME.onrender.com/scrape" \
  -H "Content-Type: application/json" \
  -d '{"query":"dentists","location":"Vellore, Tamil Nadu","max_results":20,"scroll_pause":2}' \
  --output google_maps_businesses.xlsx
```

## Important

The API is intentionally simple and has no authentication. If you make the URL public, anyone who discovers it can trigger scraping jobs. Add authentication/rate limiting before using it as a public production service.

The service also performs browser automation against Google Maps. Make sure your use complies with applicable terms, laws, and policies.
