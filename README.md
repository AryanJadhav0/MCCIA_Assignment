# Supplier Intelligence Dashboard

A comprehensive, interactive dashboard for analyzing supplier performance, evaluating financial impact, and generating automated negotiation briefs.

## Features

- **Dashboard Overview**: High-level KPIs tracking total billed amounts, headline losses, excess billing gaps, rejections, and returns. Includes interactive visualizations (Chart.js) for loss composition and monthly billing gap trends.
- **Supplier Scorecard**: A searchable, sortable table ranking 34 suppliers based on a composite performance score (0-100), categorizing them into `POOR` and `ACCEPTABLE` tiers.
- **Automated Negotiation Briefs**: Self-generating, printable strategy briefs for underperforming suppliers. Each brief outlines financial impact summaries, operational metrics, and a step-by-step required action plan for the supplier.
- **100% Client-Side**: The entire dataset is securely embedded directly within the application, ensuring lightning-fast load times, complete data portability, and zero backend dependencies. 

## Deployment

This dashboard is built as a self-contained static HTML file (`index.html`), making it incredibly easy to host. 

It is designed to be instantly deployable to platforms like [Vercel](https://vercel.com/) or GitHub Pages with **zero configuration** required.

## Local Development

To view the dashboard locally, simply double-click the `index.html` file to open it in any modern web browser.

Alternatively, you can serve it via a local HTTP server:

```bash
python -m http.server 8000
```
Then navigate to `http://localhost:8000/`.
