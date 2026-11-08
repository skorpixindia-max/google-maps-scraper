# Google Maps Business Scraper

An interactive Node.js command-line tool that searches Google Maps for business listings and exports selected profile details to CSV and Excel. Configure locations, categories, filters, and scrape limits through the prompts, then review the results in the `results/` directory.

> **Important:** Use this tool only where you have a lawful basis to collect and use the data. Review Google's terms and policies, respect applicable privacy and data-protection laws, and keep requests at a responsible rate. Google Maps may limit or block automated access.

## Quick start

**Requirements:** Node.js 18 or later and npm.

```bash
git clone https://github.com/skorpixindia-max/google-maps-scraper.git
cd google-maps-scraper
npm install
node main.js
```

Follow the interactive prompts to choose locations and categories, set phone and website filters, enable or disable geo-grid search, and choose a listing target and time limit. Confirm the summary to begin.

## What it does

- Searches selected cities and areas for business categories
- Includes location and category choices across several countries, with prompts for custom countries, cities, and areas
- Offers an optional geo-grid search for areas with known coordinates
- Filters listings by phone-number and website availability
- Skips closed listings and configured large chains, and avoids duplicate profiles
- Saves progress so a run can be continued or selected work can be re-scraped
- Writes combined and per-city/category results as CSV and Excel workbooks

Search results depend on Google Maps availability, the selected area and category, and the configured limits. The listing target is a maximum, not a guarantee of how many results will be found.

## How to run

From the project directory, run the main script directly:

```bash
node main.js
```

You can also use the npm script, which runs the same command:

```bash
npm start
```

At launch, choose one of the available modes:

1. **New scrape** — start a fresh run.
2. **Continue** — load saved progress and skip completed city/category work.
3. **Re-scrape** — select work to run again.

Then select locations and categories, apply filters, choose whether to use geo-grid search, and set a target and time limit. Review the summary and enter `Y` to start.

The CLI is interactive and expects a terminal with standard input available.

## Configuration

The current scraper is configured through its interactive prompts. Although the repository contains an `INPUT.json` example, `main.js` does not currently read that file; editing it will not change the run. The default example lists Chennai, several business categories, and a target of 500 listings, but select the desired settings in the CLI.

To add a custom country, city, area, or category for a run, use the corresponding prompts. Geo-grid search uses built-in coordinates for supported areas; areas without a matching coordinate use text search instead.

## Results and saved state

Generated files are saved under `results/`:

- `ALL_RESULTS_COMBINED.csv` and `ALL_RESULTS_COMBINED.xlsx` — combined results
- City/category/date-named `.csv` and `.xlsx` files — per-work-unit exports

The exports can include city, region, country, category, business name, phone, rating, review count, website, coordinates, Google Maps URL, and scrape date. Some fields may be empty when they are not available in a listing.

The scraper also maintains local state in `resume_progress.json` and `scrape_history.log`. Keep these files if you want to continue a run. Back them up before deleting or replacing them; they may contain collected profile data and progress information.

## Development

Install dependencies and run the tool with Node.js 18 or later:

```bash
npm install
node main.js
```

The main implementation is in `main.js`. The project uses Playwright and Crawlee for browser automation, Apify's Actor APIs for dataset output, and ExcelJS for workbook generation.

## Responsible use

- Confirm that your intended collection and use comply with applicable laws, privacy requirements, and Google's terms and policies.
- Collect only information you are authorized to use, and handle personal or business data appropriately.
- Avoid excessive requests. Automated access may be interrupted or restricted by Google.
- Do not use collected information for spam, harassment, or other unlawful purposes.

## License

`package.json` declares the project license as Apache-2.0. Review the repository's license files and dependencies before redistributing or using the project.
