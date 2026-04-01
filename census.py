"""Fetch Census ACS data and TIGERweb centroids, merge by GEOID."""

import hashlib
import json
import logging
from pathlib import Path

import requests

from market_config import CENSUS_API_KEY

log = logging.getLogger(__name__)

ACS_BASE = "https://api.census.gov/data/{year}/acs/acs5"

# B25007: Tenure by Age of Householder (units)
# B25032: Tenure by Units in Structure (units — for sub-50 ratio)
# B25033: Total Population by Tenure by Units in Structure (people)
ACS_FIELDS = (
    "?get=B25007_001E,B25007_012E,B25007_013E,B25007_014E,"
    "B25032_013E,B25032_018E,B25032_019E,B25032_020E,B25032_021E,"
    "B25033_008E,B25033_009E,B25033_010E,B25033_011E,"
    "NAME"
)

TIGERWEB_URL = (
    "https://tigerweb.geo.census.gov/arcgis/rest/services/"
    "TIGERweb/tigerWMS_Census2020/MapServer/8/query"
)

TIMEOUT = 30

# File-based cache directory
CACHE_DIR = Path(__file__).parent / ".cache"


def _cache_path(key: str) -> Path:
    """Return a cache file path for the given key."""
    CACHE_DIR.mkdir(exist_ok=True)
    safe_key = hashlib.md5(key.encode()).hexdigest()
    return CACHE_DIR / f"{safe_key}.json"


def _cache_get(key: str):
    """Return cached data or None."""
    path = _cache_path(key)
    if path.exists():
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return None


def _cache_set(key: str, data):
    """Write data to cache."""
    path = _cache_path(key)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f)


def _safe_int(val):
    """Convert Census value to int, treating suppressed/null as 0."""
    if val is None:
        return 0
    try:
        n = int(val)
        return 0 if n < 0 else n  # negative sentinel = suppressed
    except (ValueError, TypeError):
        return 0


def _parse_acs_rows(rows: list[list]) -> list[dict]:
    """Parse ACS JSON array-of-arrays into list of dicts."""
    header = rows[0]
    col = {name: i for i, name in enumerate(header)}

    # Support both ucgid and state/county/tract/block group formats
    has_ucgid = "ucgid" in col

    results = []
    for row in rows[1:]:
        if has_ucgid:
            ucgid = row[col["ucgid"]]
            geoid = ucgid.split("US")[-1] if "US" in ucgid else ucgid
        else:
            # Older &for=block+group format
            geoid = row[col["state"]] + row[col["county"]] + row[col["tract"]] + row[col["block group"]]

        # B25032 — unit counts by structure size
        renter_units_b25032 = _safe_int(row[col["B25032_013E"]])
        renter_5to9 = _safe_int(row[col["B25032_018E"]])
        renter_10to19 = _safe_int(row[col["B25032_019E"]])
        renter_20to49 = _safe_int(row[col["B25032_020E"]])
        renter_50plus = _safe_int(row[col["B25032_021E"]])
        renter_units_sub50 = max(0, renter_units_b25032 - renter_50plus)  # clamp to 0
        # 5+ sub-50 units (for allocating 5+ population)
        fiveplus_sub50 = renter_5to9 + renter_10to19 + renter_20to49

        # B25033 — renter population by structure size
        renter_pop_total = _safe_int(row[col["B25033_008E"]])
        renter_pop_1unit = _safe_int(row[col["B25033_009E"]])
        renter_pop_2to4 = _safe_int(row[col["B25033_010E"]])
        renter_pop_5plus = _safe_int(row[col["B25033_011E"]])

        results.append({
            "geoid": geoid,
            "name": row[col["NAME"]],
            "total_units": _safe_int(row[col["B25007_001E"]]),
            "renter_total": _safe_int(row[col["B25007_012E"]]),
            "renter_15_24": _safe_int(row[col["B25007_013E"]]),
            "renter_25_34": _safe_int(row[col["B25007_014E"]]),
            # B25032 unit counts
            "renter_units_b25032": renter_units_b25032,
            "renter_units_50plus": renter_50plus,
            "renter_units_sub50": renter_units_sub50,
            "fiveplus_sub50_units": fiveplus_sub50,
            # B25033 population
            "renter_pop_total": renter_pop_total,
            "renter_pop_1unit": renter_pop_1unit,
            "renter_pop_2to4": renter_pop_2to4,
            "renter_pop_5plus": renter_pop_5plus,
        })
    return results


def fetch_acs_data(year: int, county_fips: list[str]) -> list[dict]:
    """Fetch B25007/B25032/B25033 data for the given county block groups."""
    cache_key = f"acs_{year}_{'_'.join(sorted(county_fips))}"
    cached = _cache_get(cache_key)
    if cached is not None:
        log.info("Using cached ACS data for %d (%d block groups)", year, len(cached))
        return cached

    results = []
    for fips in county_fips:
        url = (
            ACS_BASE.format(year=year)
            + ACS_FIELDS
            + f"&ucgid=pseudo(0500000US{fips}$1500000)"
        )
        if CENSUS_API_KEY:
            url += f"&key={CENSUS_API_KEY}"
        resp = requests.get(url, timeout=TIMEOUT)
        resp.raise_for_status()
        results.extend(_parse_acs_rows(resp.json()))
    log.info("Fetched %d block groups from ACS %d", len(results), year)

    _cache_set(cache_key, results)
    return results


def fetch_centroids(states: dict[str, list[str]]) -> dict[str, tuple[float, float]]:
    """Fetch block-group centroids from TIGERweb for the given state/counties.

    Args:
        states: dict mapping state FIPS -> list of 3-digit county codes
                e.g. {"34": ["023", "035"], "42": ["027"]}
    """
    cache_key = f"centroids_{'_'.join(f'{st}:{','.join(sorted(cs))}' for st, cs in sorted(states.items()))}"
    cached = _cache_get(cache_key)
    if cached is not None:
        centroids = {k: tuple(v) for k, v in cached.items()}
        log.info("Using cached centroids (%d block groups)", len(centroids))
        return centroids

    centroids: dict[str, tuple[float, float]] = {}

    MAX_PAGES = 20  # Guard against infinite pagination loops

    for state_fips, county_codes in states.items():
        for county in county_codes:
            offset = 0
            page = 0
            while page < MAX_PAGES:
                params = {
                    "where": f"STATE='{state_fips}' AND COUNTY='{county}'",
                    "outFields": "GEOID,CENTLAT,CENTLON",
                    "returnGeometry": "false",
                    "f": "json",
                    "resultRecordCount": 1000,
                    "resultOffset": offset,
                }
                resp = requests.get(TIGERWEB_URL, params=params, timeout=TIMEOUT)
                resp.raise_for_status()
                data = resp.json()

                features = data.get("features", [])
                for feat in features:
                    attrs = feat["attributes"]
                    geoid = attrs["GEOID"]
                    try:
                        lat = float(attrs["CENTLAT"])
                        lon = float(attrs["CENTLON"])
                        centroids[geoid] = (lat, lon)
                    except (ValueError, KeyError):
                        log.warning("Bad TIGERweb row for GEOID %s", geoid)

                # Check if there are more pages
                if data.get("exceededTransferLimit", False) and len(features) > 0:
                    offset += len(features)
                    page += 1
                    log.info("TIGERweb pagination: fetching page %d (offset %d) for state %s county %s",
                             page, offset, state_fips, county)
                else:
                    break

    log.info("Loaded %d centroids from TIGERweb", len(centroids))

    # Cache centroids (convert tuples to lists for JSON)
    _cache_set(cache_key, {k: list(v) for k, v in centroids.items()})
    return centroids


class MergeError(Exception):
    """Raised when too many block groups fail to match centroids."""
    pass


def merge_data(
    acs_data: list[dict],
    centroids: dict[str, tuple[float, float]],
) -> list[dict]:
    """Join ACS records with centroid coordinates by GEOID.

    Raises MergeError if >10% of block groups have no centroid match.
    """
    if not acs_data:
        raise MergeError("No ACS data to merge — Census API returned no block groups")

    merged = []
    skipped = 0
    skipped_geoids = []
    for rec in acs_data:
        coords = centroids.get(rec["geoid"])
        if coords is None:
            skipped += 1
            if len(skipped_geoids) < 10:
                skipped_geoids.append(rec["geoid"])
            continue
        merged.append({**rec, "lat": coords[0], "lon": coords[1]})

    skip_rate = skipped / len(acs_data) if acs_data else 0

    if skipped:
        log.warning(
            "Skipped %d/%d block groups with no centroid match (%.1f%%). Examples: %s",
            skipped, len(acs_data), skip_rate * 100, skipped_geoids[:5]
        )

    if skip_rate > 0.10:
        raise MergeError(
            f"Too many block groups unmatched: {skipped}/{len(acs_data)} ({skip_rate:.0%}). "
            "This may indicate a TIGERweb API issue or geography vintage mismatch."
        )

    if not merged:
        raise MergeError(
            "No block groups matched between ACS data and centroids. "
            "Check that the county FIPS codes are correct."
        )

    return merged
