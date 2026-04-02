"""Shadow Market Analysis — configurable for any university market."""

import asyncio
import json
import logging
import os
from pathlib import Path

import uvicorn
from fastapi import FastAPI, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response

from market_config import load_market, list_markets, create_market, CENSUS_API_KEY
from census import fetch_acs_data, fetch_centroids, merge_data, MergeError
from analysis import analyze, nearest_campus, assign_ring
from report import build_report, build_master_report

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

if not CENSUS_API_KEY:
    log.warning(
        "CENSUS_API_KEY not set. The Census API may rate-limit requests. "
        "Set the CENSUS_API_KEY environment variable to avoid this."
    )

app = FastAPI(title="Shadow Market Analysis")

STATIC_DIR = Path(__file__).parent / "static"
DEBUG_MODE = os.environ.get("DEBUG", "").lower() in ("1", "true", "yes")
DEBUG_OUTPUT = Path(__file__).parent / "debug_last_report.xlsx"


@app.get("/", response_class=HTMLResponse)
async def index():
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")


@app.get("/markets")
async def get_markets():
    """Return list of available market configurations."""
    return JSONResponse(content=list_markets())


@app.post("/markets/create")
async def create_market_endpoint(request: Request):
    """Create a new market config from Census API URL(s) + campus coordinates."""
    try:
        body = await request.json()
    except Exception:
        raise HTTPException(400, "Invalid JSON body")

    name = body.get("name")
    campuses = body.get("campuses", {})

    if not name:
        raise HTTPException(400, "Market name is required")
    if not campuses:
        raise HTTPException(400, "At least one campus with coordinates is required")

    try:
        config = await asyncio.to_thread(
            create_market,
            name=name,
            campuses=campuses,
            ring_miles=body.get("ring_miles"),
            years=body.get("years"),
        )
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    except Exception as exc:
        log.exception("Market creation failed")
        raise HTTPException(502, str(exc))

    return JSONResponse(content={
        "short_name": config.short_name,
        "name": config.name,
        "state_fips": config.state_fips,
        "county_fips": config.county_fips,
        "county_names": config.county_names,
        "campuses": list(config.campuses.keys()),
    })


@app.get("/markets/{short_name}")
async def get_market_detail(short_name: str):
    """Return full market config including campus coordinates."""
    try:
        config = load_market(short_name)
    except (FileNotFoundError, ValueError) as exc:
        raise HTTPException(404, str(exc))

    return JSONResponse(content={
        "short_name": config.short_name,
        "name": config.name,
        "county_names": config.county_names,
        "campuses": {k: list(v) for k, v in config.campuses.items()},
        "ring_miles": config.ring_miles,
        "ring_labels": config.ring_labels,
        "years": config.years,
    })


def _fetch_map_points(config):
    """Fetch centroids and compute ring assignments — no ACS data needed."""
    centroids = fetch_centroids(config.states)
    points = []
    for geoid, (lat, lon) in centroids.items():
        campus_name, dist = nearest_campus(lat, lon, config.campuses)
        ring = assign_ring(dist, config.ring_miles, config.ring_labels)
        if ring is not None:
            points.append({"lat": lat, "lon": lon, "ring": ring, "dist": round(dist, 2)})
    return points


@app.get("/map-data/{short_name}")
async def get_map_data(short_name: str):
    """Return block group centroids with ring assignments for map overlay."""
    try:
        config = load_market(short_name)
    except (FileNotFoundError, ValueError) as exc:
        raise HTTPException(404, str(exc))

    try:
        points = await asyncio.to_thread(_fetch_map_points, config)
    except Exception as exc:
        log.exception("Failed to fetch map data")
        raise HTTPException(502, str(exc))

    return JSONResponse(content=points)


def _run_single_report(config, year: int) -> tuple[dict, bytes]:
    """Synchronous report generation — run in thread pool."""
    acs = fetch_acs_data(year, config.county_fips)
    centroids = fetch_centroids(config.states)
    merged = merge_data(acs, centroids)
    result = analyze(merged, config.campuses, config.ring_miles, config.ring_labels)

    # Guard against empty results
    if result["total"]["block_groups"] == 0:
        raise ValueError(
            "No block groups fell within the distance rings. "
            "Check that campus coordinates are near the specified counties."
        )

    xlsx = build_report(result, year, config)
    return result, xlsx


def _run_master_report(config) -> tuple[dict, bytes]:
    """Synchronous master report generation — run in thread pool."""
    centroids = fetch_centroids(config.states)
    yearly_results = {}
    for year in config.years:
        acs = fetch_acs_data(year, config.county_fips)
        merged = merge_data(acs, centroids)
        yearly_results[year] = analyze(merged, config.campuses, config.ring_miles, config.ring_labels)

    # Guard against empty results
    latest = yearly_results[max(config.years)]
    if latest["total"]["block_groups"] == 0:
        raise ValueError(
            "No block groups fell within the distance rings. "
            "Check that campus coordinates are near the specified counties."
        )

    xlsx = build_master_report(yearly_results, config)
    return yearly_results, xlsx


@app.post("/report")
async def generate_report(year: int = Form(2024), market: str = Form("rutgers_nb")):
    try:
        config = load_market(market)
    except (FileNotFoundError, ValueError) as exc:
        raise HTTPException(400, str(exc))

    if year not in config.years:
        raise HTTPException(400, f"Year must be one of {config.years}")

    try:
        result, xlsx = await asyncio.to_thread(_run_single_report, config, year)
    except (MergeError, ValueError) as exc:
        raise HTTPException(400, str(exc))
    except Exception as exc:
        log.exception("Report generation failed")
        raise HTTPException(status_code=502, detail=str(exc))

    if DEBUG_MODE:
        DEBUG_OUTPUT.write_bytes(xlsx)

    total = result["total"]
    summary = (
        f"{total['block_groups']} block groups, "
        f"{total['renter_15_24']} renters aged 15-24 within {config.ring_miles[-1]} miles"
    )

    return Response(
        content=xlsx,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename=shadow_market_{config.short_name}_{year}.xlsx",
            "X-Summary": summary,
        },
    )


@app.post("/master-report")
async def generate_master_report(market: str = Form("rutgers_nb")):
    try:
        config = load_market(market)
    except (FileNotFoundError, ValueError) as exc:
        raise HTTPException(400, str(exc))

    try:
        yearly_results, xlsx = await asyncio.to_thread(_run_master_report, config)
    except (MergeError, ValueError) as exc:
        raise HTTPException(400, str(exc))
    except Exception as exc:
        log.exception("Master report generation failed")
        raise HTTPException(status_code=502, detail=str(exc))

    if DEBUG_MODE:
        DEBUG_OUTPUT.write_bytes(xlsx)

    latest = yearly_results[max(config.years)]["total"]
    year_range = f"{min(config.years)}-{max(config.years)}"
    summary = (
        f"{latest['block_groups']} block groups, "
        f"{latest['renter_15_24']} renters aged 15-24 (latest year) across {year_range}"
    )

    return Response(
        content=xlsx,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f"attachment; filename=shadow_market_{config.short_name}_master.xlsx",
            "X-Summary": summary,
        },
    )


if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8004)
