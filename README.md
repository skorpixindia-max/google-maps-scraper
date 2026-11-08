# Google Maps Business Scraper

A Node.js scraper for collecting business listings from Google Maps using Playwright and Crawlee. It supports city and category targeting, geo-grid search, deduplication, resume logic, and exports to CSV and Excel.

## Features

- Search Google Maps by city, region, and category
- Multi-country and multi-area scraping
- Geo-grid location search to improve discovery coverage
- Filters for phone number and website availability
- Duplicate protection and resume support
- Exports results as CSV and Excel (.xlsx)
- Saves progress and scrape history to local files

## Tech Stack

- Node.js 18+
- Playwright
- Crawlee
- Apify
- ExcelJS

## Project Structure

- `main.js` — scraper logic and interactive menu flow
- `INPUT.json` — default cities, categories, and target count
- `results/` — exported CSV and Excel files
- `resume_progress.json` — resume state snapshot
- `scrape_history.log` — execution history
- `Dockerfile` — containerized runtime configuration

## Installation

```bash
npm install
```

## Configuration

Edit `INPUT.json` to set the search inputs:

```json
{
  "cities": ["Chennai"],
  "businessCategories": [
    "dentist",
    "dental clinic",
    "skin clinic",
    "gym"
  ],
  "targetListings": 500
}
```

The scraper will prompt you interactively for:
- country / region / city / area selection
- business categories
- phone and website filters
- geo-grid enable/disable
- time and target settings

## Running the scraper

```bash
npm start
```

When started, the program runs an interactive menu and guides you through the scrape flow.

## Output Files

Results are generated in the `results/` directory, including:

- `ALL_RESULTS_COMBINED.csv`
- `ALL_RESULTS_COMBINED.xlsx`
- per-city and per-category exports
- progress/resume state files

Each output includes fields such as:
- City
- State / Region
- Country
- Category
- Business Name
- Phone
- Rating
- Reviews
- Website
- Latitude
- Longitude
- Google Maps URL
- Scraped Date

## Notes

- This project is designed for local data collection and research-oriented workflows.
- Google Maps may rate-limit or block automation; the scraper includes retry, delay, and block detection logic.
- Use it responsibly and ensure compliance with local laws, platform terms, and website usage policies.

## License

This project is licensed under the Apache 2.0 License.

## Example

A typical run can scrape multiple categories across selected urban areas and save structured business data into Excel and CSV files for downstream analysis.
