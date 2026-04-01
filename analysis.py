"""Haversine distance calculation, ring assignment, and metric aggregation."""

import math


def haversine_miles(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Great-circle distance in miles between two lat/lon points."""
    R = 3958.8  # Earth radius in miles
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(math.radians(lat1))
        * math.cos(math.radians(lat2))
        * math.sin(dlon / 2) ** 2
    )
    return R * 2 * math.asin(min(1.0, math.sqrt(a)))


def nearest_campus(lat: float, lon: float, campuses: dict[str, tuple[float, float]]) -> tuple[str, float]:
    """Return (campus_name, distance_miles) for the nearest campus."""
    best_name = ""
    best_dist = float("inf")
    for name, (clat, clon) in campuses.items():
        d = haversine_miles(clat, clon, lat, lon)
        if d < best_dist:
            best_dist = d
            best_name = name
    return best_name, best_dist


def assign_ring(distance: float, ring_miles: list[float], ring_labels: list[str]) -> str | None:
    """Return ring label for a given distance, or None if beyond all rings."""
    for boundary, label in zip(ring_miles, ring_labels):
        if distance <= boundary:
            return label
    return None


def _compute_shadow_market(rec: dict) -> dict:
    """Compute per-block-group shadow market metrics.

    Methodology (matching reference spreadsheet):
    1. Sub-50 ratio = (total renter units - 50+ units) / total renter units
    2. Renter pop sub-50 = pop in 1-unit + pop in 2-4 + (pop in 5+ × 5+ sub-50 ratio)
       where 5+ sub-50 ratio = (5-9 + 10-19 + 20-49 units) / (5-9 + 10-19 + 20-49 + 50+ units)
       This allocates 5+ population to sub-50 using unit-count ratios, since B25033
       only provides "5 or more" as the finest category.
    3. Avg occ per sub-50 unit = pop sub-50 / units sub-50
    4. Shadow market HHs = renters 15-24 × sub-50 ratio
    5. Shadow market pop = shadow HHs × avg occ per sub-50 unit
    """
    renter_units = rec.get("renter_units_b25032", 0)
    renter_units_sub50 = rec.get("renter_units_sub50", 0)
    renter_units_50plus = rec.get("renter_units_50plus", 0)
    fiveplus_sub50_units = rec.get("fiveplus_sub50_units", 0)

    renter_pop_total = rec.get("renter_pop_total", 0)
    renter_pop_1unit = rec.get("renter_pop_1unit", 0)
    renter_pop_2to4 = rec.get("renter_pop_2to4", 0)
    renter_pop_5plus = rec.get("renter_pop_5plus", 0)

    # Sub-50 ratio: what fraction of renter units are in sub-50 buildings
    sub50_ratio = (renter_units_sub50 / renter_units) if renter_units > 0 else 0.0

    # Allocate 5+ population between sub-50 and 50+ using unit-count ratios
    fiveplus_total_units = fiveplus_sub50_units + renter_units_50plus
    if fiveplus_total_units > 0:
        fiveplus_sub50_ratio = fiveplus_sub50_units / fiveplus_total_units
    else:
        fiveplus_sub50_ratio = 0.0

    renter_pop_5plus_sub50 = renter_pop_5plus * fiveplus_sub50_ratio

    # Total sub-50 population: 1-unit + 2-4 are fully sub-50; 5+ is allocated
    renter_pop_sub50 = renter_pop_1unit + renter_pop_2to4 + renter_pop_5plus_sub50

    # Average occupancy per sub-50 renter unit
    avg_occ_sub50 = (renter_pop_sub50 / renter_units_sub50) if renter_units_sub50 > 0 else 0.0

    # Shadow market: 15-24 renters adjusted by sub-50 ratio
    renter_15_24 = rec.get("renter_15_24", 0)
    shadow_hhs = renter_15_24 * sub50_ratio
    shadow_pop = shadow_hhs * avg_occ_sub50

    return {
        "renter_units_sub50": renter_units_sub50,
        "renter_units_50plus": renter_units_50plus,
        "sub50_ratio": sub50_ratio,
        "renter_pop": renter_pop_total,
        "renter_pop_sub50": renter_pop_sub50,
        "avg_occ_sub50": avg_occ_sub50,
        "shadow_hhs": shadow_hhs,
        "shadow_pop": shadow_pop,
    }


def _empty_agg():
    return {
        "block_groups": 0,
        "total_units": 0,
        "renter_units": 0,
        "renter_15_24": 0,
        "renter_25_34": 0,
        # Shadow market aggregates
        "renter_units_sub50": 0,
        "renter_units_50plus": 0,
        "renter_pop": 0,
        "renter_pop_sub50": 0.0,
        "shadow_hhs": 0.0,
        "shadow_pop": 0.0,
    }


def _add_pcts(agg: dict):
    """Add percentage and derived fields to an aggregate dict."""
    ru = agg["renter_units"]
    tu = agg["total_units"]
    agg["pct_renter"] = (ru / tu * 100) if tu else 0.0
    agg["pct_15_24_of_renter"] = (agg["renter_15_24"] / ru * 100) if ru else 0.0
    agg["pct_15_24_of_total"] = (agg["renter_15_24"] / tu * 100) if tu else 0.0

    # Overall sub-50 ratio and avg occupancy for this ring/total
    sub50 = agg["renter_units_sub50"]
    agg["sub50_ratio"] = (sub50 / ru * 100) if ru else 0.0
    agg["avg_occ_sub50"] = (agg["renter_pop_sub50"] / sub50) if sub50 else 0.0


def analyze(
    merged_data: list[dict],
    campuses: dict[str, tuple[float, float]],
    ring_miles: list[float],
    ring_labels: list[str],
) -> dict:
    """Compute distance rings and aggregate metrics.

    Each block group is assigned to the ring of its nearest campus.
    No block group is counted more than once.

    Returns dict with keys: rings, total, detail.
    """
    rings = {label: _empty_agg() for label in ring_labels}
    detail = []

    for rec in merged_data:
        campus_name, dist = nearest_campus(rec["lat"], rec["lon"], campuses)
        ring = assign_ring(dist, ring_miles, ring_labels)
        if ring is None:
            continue

        # Per-block-group shadow market calc
        sm = _compute_shadow_market(rec)

        agg = rings[ring]
        agg["block_groups"] += 1
        agg["total_units"] += rec["total_units"]
        agg["renter_units"] += rec["renter_total"]
        agg["renter_15_24"] += rec["renter_15_24"]
        agg["renter_25_34"] += rec["renter_25_34"]
        agg["renter_units_sub50"] += sm["renter_units_sub50"]
        agg["renter_units_50plus"] += sm["renter_units_50plus"]
        agg["renter_pop"] += sm["renter_pop"]
        agg["renter_pop_sub50"] += sm["renter_pop_sub50"]
        agg["shadow_hhs"] += sm["shadow_hhs"]
        agg["shadow_pop"] += sm["shadow_pop"]

        detail.append({
            **rec,
            **sm,
            "distance_mi": round(dist, 2),
            "nearest_campus": campus_name,
            "ring": ring,
        })

    # Percentages per ring
    for agg in rings.values():
        _add_pcts(agg)

    # Total across all rings
    total = _empty_agg()
    sum_keys = (
        "block_groups", "total_units", "renter_units", "renter_15_24", "renter_25_34",
        "renter_units_sub50", "renter_units_50plus", "renter_pop", "renter_pop_sub50",
        "shadow_hhs", "shadow_pop",
    )
    for agg in rings.values():
        for key in sum_keys:
            total[key] += agg[key]
    _add_pcts(total)

    detail.sort(key=lambda r: r["distance_mi"])

    return {"rings": rings, "total": total, "detail": detail}
