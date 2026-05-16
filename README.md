# HERVeg.05 Intelligence Platform

Executive agricultural intelligence dashboard for program leaders, investors, NGOs, governments, and donor-facing teams.

## What It Includes

- Google Sheets data integration
- Executive overview with program health scoring
- Enrollment, payment, delivery, geography, pipeline, and strategic intelligence views
- ECharts visual analytics
- Leaflet map with cluster, heat, and zone modes
- Responsive desktop and mobile navigation
- Filter system for time, zone, village, VANO, package, and gender
- CSV, Excel, and PNG exports
- Accessible focus states, keyboard shortcuts, and reduced-motion support
- Enterprise-grade light and dark modes

## Project Structure

```text
.
├── index.html
├── styles.css
├── app.js
├── assets/
├── .gitignore
└── README.md
```

## Run Locally

This is a static app. You can open `index.html` directly, but a local server is recommended for consistent export behavior.

```powershell
python -m http.server 8080
```

Then open:

```text
http://localhost:8080
```

## Configuration

Google Sheets settings are defined in `app.js`:

```js
const CONFIG = {
  apiKey: "...",
  spreadsheetId: "...",
  sheets: {
    farmers: "Farmers",
    payments: "Farmers Payments",
    delivery: "Farmers Delivery"
  }
};
```

The spreadsheet must be readable by the configured API key and expose the expected sheet names and column headers.

## Deployment

The app can be deployed to GitHub Pages, Netlify, Vercel, Cloudflare Pages, or any static web host.

For GitHub Pages:

1. Push this repository to GitHub.
2. Open repository settings.
3. Enable Pages from the default branch root.
4. Visit the generated GitHub Pages URL.

## Notes

The dashboard intentionally avoids decorative visual effects and prioritizes executive comprehension, data clarity, and fast decision-making.
