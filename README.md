# Shadow Market Analysis Tool

Estimates the shadow market college-age renter population near university campuses by combining CoStar building data with Census ACS demographics. Generates styled Excel reports with distance-ring breakdowns, interactive map visualization, and multi-year trend analysis.

## Quick Start

1. **Clone the repo**
   ```
   git clone https://github.com/jacksubtextuse-ux/shadow-market-tool.git
   cd shadow-market-tool
   ```

2. **Install dependencies**
   ```
   pip install -r requirements.txt
   ```

3. **Run the server**
   ```
   python main.py
   ```

4. **Open** `http://localhost:8004` in your browser

## Using the Tool

### Census-Only Report
- Select a market from the dropdown
- Pick a year (or "Master" for all years)
- Click **Generate Census Report**
- Excel file downloads with block-group-level shadow market analysis

### CoStar Shadow Market Analysis
This is the primary analysis mode. It combines actual CoStar building data with Census demographics:

1. Select a market from the dropdown
2. Upload a **CoStar CSV file** (must have columns: Property Address, PropertyType, University, Number Of Units, Beds, Latitude, Longitude)
3. Choose **Analysis Mode**:
   - **Combined** (recommended) -- CoStar buildings (5-49 units) + Census 2-4 unit rentals + college-age demographics
   - **Beds Only** -- raw CoStar unit/bed counts by ring, no Census
4. Toggle **Include Graduates (22-24)** to expand the age window from 18-21 to 18-24
5. Click **Run CoStar Analysis**

**Filters applied automatically:**
- Buildings with 50+ units are excluded (tracked internally)
- Properties typed as "Student" are excluded (purpose-built student housing)
- Properties named "Demolished" are excluded
- Single-family rentals (1-unit) are excluded from Census data

### Create a New Market
- Click the **+ New Market** tab
- Enter the market name (e.g. "Penn State University Park")
- Add one or more campus center points (name + latitude + longitude)
- Set distance rings (defaults to 0.5, 1, 2 miles)
- Set ACS years (defaults to 2022, 2023, 2024)
- Click **Create Market**

The tool auto-detects which counties fall within your distance rings using the Census geocoder. No manual FIPS codes needed.

### Map View
- Click the **Map View** tab
- Select a market to see campus pins, distance ring circles, and block group centroid dots colored by ring

## Shadow Market Methodology (Combined Mode)

For each block group within the distance rings:

1. **Units** = CoStar buildings (5-49 units) matched to nearest block group + Census 2-4 unit renter count (B25032) for that block group
2. **Avg Occupancy** = Census average people per sub-50 renter unit (B25033 population / B25032 units)
3. **Est. Population** = Units x Avg Occupancy
4. **Renter 15-24 share** = Renters aged 15-24 / Total renters (Census B25007)
5. **Age tightening** = Pop 18-21 / Pop 15-24 from Census B01001 (or 18-24 with graduates toggle)
6. **College renter %** = Renter 15-24 share x Age tightening ratio
7. **Shadow Population** = Est. Population x College renter %

All calculations are per individual block group, then summed by distance ring.

### Census Tables Used
- **B25007** -- Tenure by Age of Householder (renter units by age 15-24, 25-34)
- **B25032** -- Tenure by Units in Structure (unit counts by building size: 2-unit, 3-4, 5-9, 10-19, 20-49, 50+)
- **B25033** -- Total Population by Tenure by Units in Structure (renter population by building size)
- **B01001** -- Sex by Age (population by age: 15-17, 18-19, 20-21, 22-24 for tightening ratio)

## Census API Key (Optional)

The Census API works without a key but may rate-limit heavy usage. To avoid this:

1. Get a free key at https://api.census.gov/data/key_signup.html
2. Copy `.env.example` to `.env` and add your key:
   ```
   CENSUS_API_KEY=your_key_here
   ```
3. Set it as an environment variable before running:
   ```
   # Windows
   set CENSUS_API_KEY=your_key_here

   # Mac/Linux
   export CENSUS_API_KEY=your_key_here
   ```

## Project Structure

```
main.py              -- FastAPI server (port 8004)
market_config.py     -- Market config loader + auto county detection
census.py            -- Census ACS + TIGERweb data fetching with caching
analysis.py          -- Distance calculation + shadow market computation (Census-only mode)
costar.py            -- CoStar CSV parsing + combined analysis (CoStar + Census)
report.py            -- Excel report generation for Census-only mode
costar_report.py     -- Excel report generation for CoStar combined mode
static/index.html    -- Web UI (Run Report, + New Market, Map View tabs)
markets/             -- Market config JSON files
.cache/              -- Auto-generated data cache (gitignored)
```

## Included Markets

- **Rutgers New Brunswick** -- 4 campuses (College Ave, Busch, Livingston, Cook/Douglass), Middlesex + Somerset County
- **Ohio State** -- Main Campus, Franklin County
- **University of North Carolina** -- Main Campus, Orange County
