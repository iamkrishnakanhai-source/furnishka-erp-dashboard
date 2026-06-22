# Furnishka ERP Migration Dashboard

Live dashboard for tracking Sales Order, Sales Invoice, and Stock Entry migration progress.

## Features
- Upload working sheet and view live progress metrics
- Filter by store, date range, SKU, customer, status
- Generate ERPNext-compatible upload files:
  - **Stock Entry** (Material Receipt) — opening stock
  - **Sales Order** — pending orders
  - **Sales Invoice** — against mapped SOs
- Inline editing for pending items
- Export/Import edits for team collaboration
- Dark mode support

## Usage
1. Open the dashboard URL
2. Upload "Item Level Invoice Data Working - Final for Review.xlsx"
3. Use filters and generate upload files as needed

## Team Collaboration
- Each person opens the same URL and uploads the same Excel file
- Edit pending rows → click "Export edits" → share JSON with team
- Others click "Import edits" → edits merge automatically
