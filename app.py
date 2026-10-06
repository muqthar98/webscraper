import os
import tempfile
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from playwright.sync_api import sync_playwright

from scraper import scrape_search, format_excel


app = FastAPI(
    title="Google Maps Lead Scraper API",
    version="1.0.0",
)


class ScrapeRequest(BaseModel):
    query: str = Field(..., min_length=1, max_length=200)
    location: str = Field(..., min_length=1, max_length=200)
    max_results: int = Field(default=100, ge=1, le=500)
    scroll_pause: float = Field(default=2.0, ge=0.5, le=10.0)


@app.get("/")
def root():
    return {
        "status": "ok",
        "service": "Google Maps Lead Scraper",
        "endpoint": "POST /scrape",
    }


@app.get("/health")
def health():
    return {"status": "healthy"}


@app.post("/scrape")
def scrape(request: ScrapeRequest):
    # A temporary directory prevents concurrent requests from overwriting
    # the same Excel file.
    temp_dir = Path(tempfile.mkdtemp(prefix="maps-scrape-"))
    output_path = temp_dir / "google_maps_businesses.xlsx"

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                ],
            )

            context = browser.new_context(
                viewport={"width": 1440, "height": 1000},
                locale="en-IN",
            )

            try:
                rows = scrape_search(
                    context,
                    request.query,
                    request.location,
                    request.max_results,
                    request.scroll_pause,
                )
                count = format_excel(output_path, rows)
            finally:
                browser.close()

        return FileResponse(
            path=output_path,
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            filename="google_maps_businesses.xlsx",
            headers={"X-Result-Count": str(count)},
        )

    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Scraping failed: {str(exc)}",
        ) from exc
