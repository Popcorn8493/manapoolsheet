# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python tool for ManaPool sellers to automate fulfillment list creation. The script connects to ManaPool and Scryfall APIs to retrieve unshipped orders and create consolidated CSV files for picking and shipping.

## Core Architecture

- **Single-file application**: `manapoolsheet.py` contains the entire application logic
- **API integrations**: 
  - ManaPool API for order data (requires email + API token authentication)
  - Scryfall API for card images (with 100ms rate limiting)
- **Data processing**: Uses pandas for CSV export and data manipulation
- **Configuration**: Environment variables via .env file, locations.json for set-to-location mapping

## Essential Commands

**Run the application:**
```bash
python manapoolsheet.py
```

**Install dependencies:**
```bash
pip install -r requirements.txt
```

**Dependencies:** requests, pandas, dotenv, python-dotenv

## Key Configuration Files

- `email.txt` and `api-key.txt`: Legacy credential files (script now prefers .env)
- `.env`: Environment variables (MANAPOOL_EMAIL, MANAPOOL_API_KEY)  
- `locations.json`: Maps set codes to physical locations (optional)
- `requirements.txt`: Python dependencies

## Application Flow

1. User selects filter (not shipped/shipped/all orders)
2. Authenticates with ManaPool API using credentials
3. Fetches all orders, filters based on user choice
4. For each order, retrieves detailed line items
5. Enriches each item with Scryfall card images
6. Exports to CSV with location data from locations.json

## Output Files

- `fulfillment_list_not_shipped.csv` / `fulfillment_list_shipped.csv` / `fulfillment_list_all.csv`
- `order_retrieval_log.txt`: Application logs

## Important Implementation Details

- Rate limiting: 100ms delay between Scryfall API calls
- Error handling: Continues processing even if individual orders/images fail
- Sorting: Final CSV sorted by set, then card name
- Location mapping: Uses uppercase set codes for consistent matching