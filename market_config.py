"""Load and validate market configuration from JSON files."""

import json
import logging
import math
import os
import re
from dataclasses import dataclass, field
from pathlib import Path

import requests

log = logging.getLogger(__name__)

MARKETS_DIR = Path(__file__).parent / "markets"

# Census API key — set CENSUS_API_KEY env var to avoid rate limits
CENSUS_API_KEY = os.environ.get("CENSUS_API_KEY", "")

# Safe name pattern for market short_names (prevents path traversal)
_SAFE_NAME_RE = re.compile(r'^[a-z0-9_]+$')

GEOCODER_URL = (
    "https://geocoding.geo.census.gov/geocoder/geographies/coordinates"
)


def _validate_market_name(name: str) -> None:
    """Ensure market name is safe for filesystem use (no path traversal)."""
    if not _SAFE_NAME_RE.match(name):
        raise ValueError(f"Invalid market name '{name}'. Use only lowercase letters, numbers, and underscores.")


def _ring_labels(ring_miles: list[float]) -> list[str]:
    """Derive ring labels from mile boundaries.

    [0.5, 1.0, 2.0] -> ["0-0.5mi", "0.5-1mi", "1-2mi"]
    """
    labels = []
    prev = 0
    for m in ring_miles:
        lo = int(prev) if prev == int(prev) else prev
        hi = int(m) if m == int(m) else m
        labels.append(f"{lo}-{hi}mi")
        prev = m
    return labels


def _county_codes(county_fips: list[str]) -> list[str]:
    """Extract 3-digit county codes from full FIPS (e.g. '34023' -> '023')."""
    return [fips[-3:] for fips in county_fips]


def _validate_ring_miles(ring_miles: list[float]) -> None:
    """Ensure ring_miles are positive and in ascending order."""
    if not ring_miles:
        raise ValueError("ring_miles must not be empty.")
    for m in ring_miles:
        if m <= 0:
            raise ValueError(f"Ring distance must be positive, got {m}.")
    for i in range(1, len(ring_miles)):
        if ring_miles[i] <= ring_miles[i - 1]:
            raise ValueError(
                f"ring_miles must be in ascending order, got {ring_miles}."
            )


def _validate_coordinates(lat: float, lon: float) -> None:
    """Validate lat/lon are within valid ranges."""
    if not (-90 <= lat <= 90):
        raise ValueError(f"Invalid latitude {lat}. Must be between -90 and 90.")
    if not (-180 <= lon <= 180):
        raise ValueError(f"Invalid longitude {lon}. Must be between -180 and 180.")


@dataclass
class MarketConfig:
    name: str
    short_name: str
    state_fips: str
    county_fips: list[str]
    county_names: list[str]
    campuses: dict[str, tuple[float, float]]
    ring_miles: list[float]
    years: list[int]

    # Derived fields
    ring_labels: list[str] = field(default_factory=list)
    county_codes: list[str] = field(default_factory=list)
    # Group counties by state for multi-state support
    states: dict[str, list[str]] = field(default_factory=dict)

    def __post_init__(self):
        if not self.ring_labels:
            self.ring_labels = _ring_labels(self.ring_miles)
        if not self.county_codes:
            self.county_codes = _county_codes(self.county_fips)
        if not self.states:
            states: dict[str, list[str]] = {}
            for fips in self.county_fips:
                st = fips[:2]
                county_code = fips[-3:]
                states.setdefault(st, []).append(county_code)
            self.states = states


def load_market(name: str) -> MarketConfig:
    """Load a market config from markets/{name}.json."""
    _validate_market_name(name)

    path = MARKETS_DIR / f"{name}.json"
    if not path.exists():
        raise FileNotFoundError(f"Market config not found: {name}")

    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    campuses = {k: tuple(v) for k, v in data["campuses"].items()}

    return MarketConfig(
        name=data["name"],
        short_name=data["short_name"],
        state_fips=data["state_fips"],
        county_fips=data["county_fips"],
        county_names=data["county_names"],
        campuses=campuses,
        ring_miles=data["ring_miles"],
        years=data["years"],
    )


def list_markets() -> list[dict]:
    """Return list of available markets with name and short_name."""
    markets = []
    for path in sorted(MARKETS_DIR.glob("*.json")):
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            markets.append({
                "short_name": data["short_name"],
                "name": data["name"],
                "years": data["years"],
                "county_names": data["county_names"],
                "campuses": list(data["campuses"].keys()),
                "ring_miles": data["ring_miles"],
            })
        except Exception:
            log.warning("Skipping invalid market config: %s", path.name)
    return markets


# ---------------------------------------------------------------------------
# Auto-detect counties from campus coordinates
# ---------------------------------------------------------------------------

def _geocode_county(lat: float, lon: float) -> tuple[str, str] | None:
    """Use Census geocoder to find the county FIPS for a lat/lon point.

    Returns (state_county_fips, county_name) e.g. ("34023", "Middlesex")
    or None if geocoding fails.
    """
    try:
        params = {
            "x": lon,
            "y": lat,
            "benchmark": "Public_AR_Current",
            "vintage": "Census2020_Current",
            "format": "json",
        }
        resp = requests.get(GEOCODER_URL, params=params, timeout=15)
        resp.raise_for_status()
        data = resp.json()

        geographies = data.get("result", {}).get("geographies", {})
        counties = geographies.get("Counties", [])
        if counties:
            c = counties[0]
            state = c.get("STATE", "")
            county = c.get("COUNTY", "")
            name = c.get("NAME", "").replace(" County", "").strip()
            if state and county:
                return (state + county, name)
    except Exception:
        log.warning("Geocoder failed for (%s, %s)", lat, lon)
    return None


def _nearby_county_fips(lat: float, lon: float, radius_miles: float) -> list[tuple[str, str]]:
    """Find counties within radius of a point by sampling cardinal directions.

    Returns list of (fips, county_name) tuples for unique counties found.
    """
    results = {}

    # Check the center point
    center = _geocode_county(lat, lon)
    if center:
        results[center[0]] = center[1]

    # Sample points at the max ring radius in 8 directions
    # 1 degree lat ~ 69 miles, 1 degree lon ~ 69 * cos(lat) miles
    lat_offset = radius_miles / 69.0
    lon_offset = radius_miles / (69.0 * math.cos(math.radians(lat)))

    offsets = [
        (lat_offset, 0),             # N
        (-lat_offset, 0),            # S
        (0, lon_offset),             # E
        (0, -lon_offset),            # W
        (lat_offset, lon_offset),    # NE
        (lat_offset, -lon_offset),   # NW
        (-lat_offset, lon_offset),   # SE
        (-lat_offset, -lon_offset),  # SW
    ]

    for dlat, dlon in offsets:
        result = _geocode_county(lat + dlat, lon + dlon)
        if result and result[0] not in results:
            results[result[0]] = result[1]
            log.info("Found nearby county: %s (%s) from offset (%.3f, %.3f)",
                     result[1], result[0], dlat, dlon)

    return [(fips, name) for fips, name in results.items()]


def detect_counties(
    campuses: dict[str, tuple[float, float]],
    max_ring_miles: float,
) -> tuple[list[str], list[str]]:
    """Auto-detect all counties that could contain block groups within the
    ring radius of any campus center point.

    Returns (county_fips_list, county_names_list).
    """
    all_counties: dict[str, str] = {}  # fips -> name

    for campus_name, (lat, lon) in campuses.items():
        log.info("Detecting counties near %s (%.4f, %.4f, radius %.1f mi)...",
                 campus_name, lat, lon, max_ring_miles)
        nearby = _nearby_county_fips(lat, lon, max_ring_miles)
        for fips, name in nearby:
            if fips not in all_counties:
                all_counties[fips] = name
                log.info("  -> %s County (%s)", name, fips)

    if not all_counties:
        raise ValueError(
            "Could not detect any counties from campus coordinates. "
            "Check that the coordinates are within the United States."
        )

    # Sort by FIPS for deterministic order
    sorted_fips = sorted(all_counties.keys())
    county_fips = sorted_fips
    county_names = [all_counties[f] for f in sorted_fips]

    return county_fips, county_names


# ---------------------------------------------------------------------------
# Create new market configs
# ---------------------------------------------------------------------------

def create_market(
    name: str,
    campuses: dict[str, list[float]],
    ring_miles: list[float] | None = None,
    years: list[int] | None = None,
) -> MarketConfig:
    """Auto-detect counties from campus coordinates and save a new market config.

    Returns the created MarketConfig.
    """
    # Validate campus coordinates
    for campus_name, coords in campuses.items():
        if len(coords) != 2:
            raise ValueError(f"Campus '{campus_name}' must have exactly [lat, lon]")
        _validate_coordinates(coords[0], coords[1])

    # Defaults
    if ring_miles is None:
        ring_miles = [0.5, 1.0, 2.0]
    if years is None:
        years = [2022, 2023, 2024]

    _validate_ring_miles(ring_miles)

    if len(years) > 10:
        raise ValueError("Maximum 10 years allowed per market config")

    # Convert campus coords to tuples
    campuses_tuples = {k: tuple(v) for k, v in campuses.items()}

    # Auto-detect counties from campus coordinates + max ring distance
    max_ring = max(ring_miles)
    county_fips, county_names = detect_counties(campuses_tuples, max_ring)

    # Derive state from first FIPS
    state_fips = county_fips[0][:2]

    # Generate short_name from market name
    short_name = re.sub(r'[^a-z0-9]+', '_', name.lower()).strip('_')
    _validate_market_name(short_name)

    # Check for existing config — prevent silent overwrite
    path = MARKETS_DIR / f"{short_name}.json"
    if path.exists():
        raise ValueError(
            f"Market config '{short_name}' already exists. "
            "Delete it first or choose a different name."
        )

    config = MarketConfig(
        name=name,
        short_name=short_name,
        state_fips=state_fips,
        county_fips=county_fips,
        county_names=county_names,
        campuses=campuses_tuples,
        ring_miles=ring_miles,
        years=years,
    )

    # Save to JSON
    save_data = {
        "name": config.name,
        "short_name": config.short_name,
        "state_fips": config.state_fips,
        "county_fips": config.county_fips,
        "county_names": config.county_names,
        "campuses": {k: list(v) for k, v in config.campuses.items()},
        "ring_miles": config.ring_miles,
        "years": config.years,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(save_data, f, indent=2)
    log.info("Saved market config: %s -> %s", config.name, path)

    return config
