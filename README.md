# Shadow Market Analysis Tool

Estimates the shadow market renter population (ages 15-24) near university campuses using Census ACS data. Generates styled Excel reports with distance-ring breakdowns and multi-year trend analysis.

## Quick Start

1. **Clone the repo**
   ```
   git clone <repo-url>
   cd shadow-market-rutgers
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

### Run a Report
- Select a market from the dropdown
- Pick a year (or "Master" for all years)
- Click **Generate Report**
- An Excel file downloads automatically

### Create a New Market
- Click the **+ New Market** tab
- Enter the market name (e.g. "Penn State University Park")
- Add one or more campus center points (name + latitude + longitude)
- Set distance rings (defaults to 0.5, 1, 2 miles)
- Set ACS years (defaults to 2022, 2023, 2024)
- Click **Create Market**

The tool auto-detects which counties fall within your distance rings using the Census geocoder. No manual FIPS codes needed.

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

## How It Works

1. **Census Data** -- Pulls ACS 5-Year tables B25007 (tenure by age), B25032 (units by structure size), and B25033 (population by structure size) at the block group level
2. **Centroids** -- Fetches block group centroid coordinates from TIGERweb
3. **Distance Rings** -- Assigns each block group to the nearest campus and a distance ring using haversine distance
4. **Shadow Market Calculation** -- For each block group:
   - Excludes 50+ unit buildings (already tracked internally)
   - Computes sub-50 unit ratio per block group
   - Allocates renter population to sub-50 buildings using Census structure-size data
   - Estimates shadow market households and population for renters aged 15-24
5. **Excel Report** -- Generates styled workbook with summary, detail, and trend sheets

## Project Structure

```
main.py              -- FastAPI server (port 8004)
market_config.py     -- Market config loader + auto county detection
census.py            -- Census ACS + TIGERweb data fetching with caching
analysis.py          -- Distance calculation + shadow market computation
report.py            -- Excel report generation (openpyxl)
static/index.html    -- Web UI
markets/             -- Market config JSON files
.cache/              -- Auto-generated data cache (gitignored)
```
