"""CoStar CSV parsing and building-level shadow market analysis."""

import csv
import logging
from analysis import haversine_miles, nearest_campus, assign_ring

log = logging.getLogger(__name__)


def parse_costar_csv(file_path: str, university_filter: str, max_units: int = 49) -> list[dict]:
    """Parse CoStar CSV, filter by university and unit count."""
    buildings = []
    with open(file_path, "r", encoding="latin-1") as f:
        reader = csv.DictReader(f)
        for row in reader:
            uni = row.get("University", "")
            if university_filter.lower() not in uni.lower():
                continue

            name = row.get("Property Name", "").strip()
            if name.lower() == "demolished":
                continue

            prop_type = row.get("PropertyType", "").strip().lower()
            if prop_type.startswith("student"):
                continue

            units = _safe_int(row.get("Number Of Units", ""))
            if units == 0 or units > max_units:
                continue

            lat = _safe_float(row.get("Latitude", ""))
            lon = _safe_float(row.get("Longitude", ""))
            if lat == 0.0 or lon == 0.0:
                continue

            buildings.append({
                "name": row.get("Property Name", "").strip(),
                "address": row.get("Property Address", "").strip(),
                "property_id": row.get("PropertyID", "").strip(),
                "units": units,
                "beds": _safe_int(row.get("Beds", "")),
                "studio": _safe_int(row.get("Number Of Studios Units", "")),
                "br1": _safe_int(row.get("Number Of 1 Bedrooms Units", "")),
                "br2": _safe_int(row.get("Number Of 2 Bedrooms Units", "")),
                "br3": _safe_int(row.get("Number Of 3 Bedrooms Units", "")),
                "br4": _safe_int(row.get("Number Of 4 Bedrooms Units", "")),
                "lat": lat,
                "lon": lon,
                "year_built": _safe_int(row.get("Year Built", "")),
            })

    log.info("Parsed %d buildings for '%s' (max %d units)", len(buildings), university_filter, max_units)
    return buildings


def parse_costar_bytes(data: bytes, university_filter: str, max_units: int = 49) -> list[dict]:
    """Parse CoStar CSV from uploaded bytes."""
    import io
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("latin-1")
    buildings = []
    reader = csv.DictReader(io.StringIO(text))
    for row in reader:
        uni = row.get("University", "")
        if university_filter.lower() not in uni.lower():
            continue

        name = row.get("Property Name", "").strip()
        if name.lower() == "demolished":
            continue

        prop_type = row.get("PropertyType", "").strip().lower()
        if prop_type.startswith("student"):
            continue

        units = _safe_int(row.get("Number Of Units", ""))
        if units == 0 or units > max_units:
            continue

        lat = _safe_float(row.get("Latitude", ""))
        lon = _safe_float(row.get("Longitude", ""))
        if lat == 0.0 or lon == 0.0:
            continue

        buildings.append({
            "name": name,
            "address": row.get("Property Address", "").strip(),
            "property_id": row.get("PropertyID", "").strip(),
            "units": units,
            "beds": _safe_int(row.get("Beds", "")),
            "studio": _safe_int(row.get("Number Of Studios Units", "")),
            "br1": _safe_int(row.get("Number Of 1 Bedrooms Units", "")),
            "br2": _safe_int(row.get("Number Of 2 Bedrooms Units", "")),
            "br3": _safe_int(row.get("Number Of 3 Bedrooms Units", "")),
            "br4": _safe_int(row.get("Number Of 4 Bedrooms Units", "")),
            "lat": lat,
            "lon": lon,
            "year_built": _safe_int(row.get("Year Built", "")),
        })

    log.info("Parsed %d buildings for '%s' (max %d units)", len(buildings), university_filter, max_units)
    return buildings


def analyze_costar(
    buildings: list[dict],
    campuses: dict[str, tuple[float, float]],
    ring_miles: list[float],
    ring_labels: list[str],
) -> dict:
    """Analyze CoStar buildings by distance ring — beds only, no occupancy."""
    rings = {label: _empty_ring() for label in ring_labels}
    detail = []

    for bldg in buildings:
        campus_name, dist = nearest_campus(bldg["lat"], bldg["lon"], campuses)
        ring = assign_ring(dist, ring_miles, ring_labels)
        if ring is None:
            continue

        agg = rings[ring]
        agg["buildings"] += 1
        agg["units"] += bldg["units"]
        agg["beds"] += bldg["beds"]
        agg["studio"] += bldg["studio"]
        agg["br1"] += bldg["br1"]
        agg["br2"] += bldg["br2"]
        agg["br3"] += bldg["br3"]
        agg["br4"] += bldg["br4"]

        detail.append({
            **bldg,
            "nearest_campus": campus_name,
            "distance_mi": round(dist, 2),
            "ring": ring,
        })

    detail.sort(key=lambda r: r["distance_mi"])

    total = _empty_ring()
    for agg in rings.values():
        for k in total:
            total[k] += agg[k]

    return {"rings": rings, "total": total, "detail": detail}


def analyze_costar_combined(
    buildings: list[dict],
    campuses: dict[str, tuple[float, float]],
    ring_miles: list[float],
    ring_labels: list[str],
    census_merged: list[dict],
    include_graduates: bool = False,
) -> dict:
    """Combined analysis: CoStar buildings (5-49 units) + Census 1-4 unit rentals.

    Per block group math:
      1. Units = CoStar buildings matched to BG + Census 1-4 unit count for BG
      2. Est Pop = units × avg_occ (Census B25033/B25032 per BG)
      3. College-age % = (18-21 or 18-24) / total_pop from Census B01001 per BG
      4. Shadow Pop = Est Pop × college-age %
    """
    bg_data = [rec for rec in census_merged if rec.get("lat") and rec.get("lon")]

    rings = {label: _empty_ring_combined() for label in ring_labels}
    detail = []

    # Track which BGs have been counted for Census 1-4 units (avoid double-counting)
    bg_census_counted = set()

    # --- Part 1: CoStar buildings (5-49 units) ---
    for bldg in buildings:
        campus_name, dist = nearest_campus(bldg["lat"], bldg["lon"], campuses)
        ring = assign_ring(dist, ring_miles, ring_labels)
        if ring is None:
            continue

        best_bg = _find_nearest_bg(bldg["lat"], bldg["lon"], bg_data)
        avg_occ = best_bg.get("avg_occ_sub50", 0) if best_bg else 0
        tighten = _college_age_tightening(best_bg, include_graduates) if best_bg else 0
        bg_renters_15_24 = best_bg.get("renter_15_24", 0) if best_bg else 0
        bg_renter_total = best_bg.get("renter_total", 0) if best_bg else 0

        # Share of renters that are 15-24 in this BG
        pct_15_24 = (bg_renters_15_24 / bg_renter_total) if bg_renter_total > 0 else 0
        # Tighten to college-age subset (18-21 or 18-24)
        college_renter_pct = pct_15_24 * tighten

        est_pop = bldg["units"] * avg_occ
        shadow_pop = est_pop * college_renter_pct

        agg = rings[ring]
        agg["buildings"] += 1
        agg["costar_units"] += bldg["units"]
        agg["costar_beds"] += bldg["beds"]
        agg["est_pop"] += est_pop
        agg["shadow_pop"] += shadow_pop

        detail.append({
            **bldg,
            "source": "CoStar",
            "nearest_campus": campus_name,
            "distance_mi": round(dist, 2),
            "ring": ring,
            "matched_bg": best_bg.get("geoid", "") if best_bg else "",
            "census_avg_occ": round(avg_occ, 2),
            "college_pct": round(college_renter_pct * 100, 1),
            "est_pop": round(est_pop, 1),
            "shadow_pop": round(shadow_pop, 1),
        })

    # --- Part 2: Census 1-4 unit rentals (per block group) ---
    for bg in bg_data:
        campus_name, dist = nearest_campus(bg["lat"], bg["lon"], campuses)
        ring = assign_ring(dist, ring_miles, ring_labels)
        if ring is None:
            continue

        geoid = bg["geoid"]
        if geoid in bg_census_counted:
            continue
        bg_census_counted.add(geoid)

        units_1to4 = bg.get("renter_units_2to4", 0)
        if units_1to4 <= 0:
            continue

        avg_occ = bg.get("avg_occ_sub50", 0)
        tighten = _college_age_tightening(bg, include_graduates)
        bg_renters_15_24 = bg.get("renter_15_24", 0)
        bg_renter_total = bg.get("renter_total", 0)
        pct_15_24 = (bg_renters_15_24 / bg_renter_total) if bg_renter_total > 0 else 0
        college_renter_pct = pct_15_24 * tighten

        est_pop = units_1to4 * avg_occ
        shadow_pop = est_pop * college_renter_pct

        agg = rings[ring]
        agg["census_2to4_bgs"] += 1
        agg["census_2to4_units"] += units_1to4
        agg["est_pop"] += est_pop
        agg["shadow_pop"] += shadow_pop

        detail.append({
            "name": bg.get("name", ""),
            "address": geoid,
            "property_id": geoid,
            "units": units_1to4,
            "beds": 0,
            "studio": 0, "br1": 0, "br2": 0, "br3": 0, "br4": 0,
            "lat": bg["lat"], "lon": bg["lon"],
            "year_built": 0,
            "source": "Census 1-4",
            "nearest_campus": campus_name,
            "distance_mi": round(dist, 2),
            "ring": ring,
            "matched_bg": geoid,
            "census_avg_occ": round(avg_occ, 2),
            "college_pct": round(college_renter_pct * 100, 1),
            "est_pop": round(est_pop, 1),
            "shadow_pop": round(shadow_pop, 1),
        })

    detail.sort(key=lambda r: r["distance_mi"])

    # Totals
    total = _empty_ring_combined()
    for agg in rings.values():
        for k in total:
            total[k] += agg[k]

    # Derived: total units and avg_occ per ring
    for agg in list(rings.values()) + [total]:
        agg["total_units"] = agg["costar_units"] + agg["census_2to4_units"]
        agg["avg_occ"] = round(agg["est_pop"] / agg["total_units"], 2) if agg["total_units"] > 0 else 0

    return {"rings": rings, "total": total, "detail": detail}


def _college_age_tightening(bg: dict, include_graduates: bool) -> float:
    """Calculate the tightening ratio to narrow B25007 renters 15-24 using B01001 age data.

    Uses B01001 to find what share of the 15-24 population is actually college-age:
      Base (18-21): (pop_18_19 + pop_20_21) / (pop_15_17 + pop_18_19 + pop_20_21 + pop_22_24)
      With graduates (18-24): (pop_18_19 + pop_20_21 + pop_22_24) / (same denominator)

    This ratio is then applied to B25007 renters_15_24 to get adjusted renter count.
    """
    pop_15_17 = bg.get("pop_15_17", 0)
    pop_18_19 = bg.get("pop_18_19", 0)
    pop_20_21 = bg.get("pop_20_21", 0)
    pop_22_24 = bg.get("pop_22_24", 0)

    pop_15_24 = pop_15_17 + pop_18_19 + pop_20_21 + pop_22_24
    if pop_15_24 <= 0:
        return 0.0

    college_pop = pop_18_19 + pop_20_21
    if include_graduates:
        college_pop += pop_22_24

    return college_pop / pop_15_24


def _find_nearest_bg(lat: float, lon: float, bg_data: list[dict]) -> dict | None:
    """Find the nearest Census block group to a building's coordinates."""
    best = None
    best_dist = float("inf")
    for bg in bg_data:
        d = haversine_miles(lat, lon, bg["lat"], bg["lon"])
        if d < best_dist:
            best_dist = d
            best = bg
    return best


def _compute_bg_occupancy(rec: dict) -> float:
    """Compute avg occupancy for sub-50 units in a Census block group.

    Same methodology as analysis.py _compute_shadow_market.
    """
    renter_units_sub50 = rec.get("renter_units_sub50", 0)
    if renter_units_sub50 <= 0:
        return 0.0

    pop_1unit = rec.get("renter_pop_1unit", 0)
    pop_2to4 = rec.get("renter_pop_2to4", 0)
    pop_5plus = rec.get("renter_pop_5plus", 0)

    fiveplus_sub50 = rec.get("fiveplus_sub50_units", 0)
    fiveplus_total = fiveplus_sub50 + rec.get("renter_units_50plus", 0)

    if fiveplus_total > 0 and pop_5plus > 0:
        ratio = fiveplus_sub50 / fiveplus_total
        pop_5plus_sub50 = pop_5plus * ratio
    else:
        pop_5plus_sub50 = 0

    pop_sub50 = pop_1unit + pop_2to4 + pop_5plus_sub50
    return pop_sub50 / renter_units_sub50


def _empty_ring() -> dict:
    return {"buildings": 0, "units": 0, "beds": 0, "studio": 0, "br1": 0, "br2": 0, "br3": 0, "br4": 0}


def _empty_ring_occ() -> dict:
    return {**_empty_ring(), "est_pop": 0, "avg_occ": 0, "shadow_units_15_24": 0, "shadow_pop_15_24": 0}


def _empty_ring_combined() -> dict:
    return {
        "buildings": 0, "costar_units": 0, "costar_beds": 0,
        "census_2to4_bgs": 0, "census_2to4_units": 0,
        "total_units": 0, "est_pop": 0, "avg_occ": 0, "shadow_pop": 0,
    }


def _safe_int(val: str) -> int:
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return 0


def _safe_float(val: str) -> float:
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0
