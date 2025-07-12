# ManaPool Order Retriever

A Python tool for sellers to automate the creation of a fulfillment list from their ManaPool account.

This script connects to the ManaPool and Scryfall APIs to retrieve orders, collect detailed data for each item, and enrich it with card images and physical storage locations. The primary goal is to create a consolidated CSV file that can be used for efficient picking and shipping of orders.

## Features

-   Fetches all or filtered (shipped/unshipped) orders from the ManaPool API.
-   For each item in an order, it retrieves detailed card information.
-   Retrieves card image URLs from the Scryfall API.
-   Maps card sets to user-defined physical locations (e.g., "Shelf 1", "Box A").
-   Securely handles API credentials using a `.env` file.
-   Exports a detailed fulfillment list to a CSV file.
-   Logs all operations to a `order_retrieval_log.txt` file.

## Setup

### 1. Prerequisites

-   Python 3.6+
-   pip (Python package installer)

### 2. Installation

1.  Clone or download this repository.
2.  Navigate to the project directory in your terminal.
3.  Install the required Python packages:
    ```bash
    pip install -r requirements.txt
    ```
    If you do not have a `requirements.txt` file, you can install the packages individually:
    ```bash
    pip install pandas python-dotenv requests
    ```

## Configuration

Before running the script, you need to configure your credentials and set locations.

### 1. API Credentials

1.  Create a file named `.env` in the root of the project directory.
2.  Add your ManaPool email and API key (retreived from https://manapool.com/seller/integrations/manapool-api) to the `.env` file as follows:

    ```
    MANAPOOL_EMAIL="your_email@example.com"
    MANAPOOL_API_KEY="your_manapool_api_key"
    ```

    **Note:** The `.env` file should never be committed to version control. It is included in the `.gitignore` file to prevent accidental exposure of your credentials.

### 2. Set Locations

1.  Create a file named `locations.json` in the root of the project directory.
2.  In this file, create a JSON object that maps your card set codes to their physical storage locations. The set codes are case-insensitive.

    **Example `locations.json`:**
    ```json
    {
      "FIN": "Shelf 1",
      "PIP": "Shelf 2",
      "FIC": "Shelf 4",
      "LCI": "Box A",
      "MKM": "Box A",
      "OTJ": "Binder 3"
    }
    ```
    If a set from an order is not found in this file, its location will be marked as "Unassigned" in the output CSV.

## Usage

1.  Open your terminal and navigate to the project directory.
2.  Run the script:
    ```bash
    python manapoolsheet.py
    ```
3.  You will be prompted to select which orders to retrieve:
    -   **1: Not Shipped (Default)** - Fetches all orders that are not yet marked as shipped.
    -   **2: Shipped Only** - Fetches all orders that have been marked as shipped.
    -   **3: All Orders** - Fetches every order in your history.
4.  The script will then fetch the order data, process it, and create a CSV file in the project directory (e.g., `fulfillment_list_not_shipped.csv`).

## Output CSV File

The generated CSV file will contain the following columns:

-   `order_id`: The unique identifier for the order.
-   `order_label`: The label assigned to the order in ManaPool.
-   `location`: The physical location of the set, based on your `locations.json` file.
-   `quantity`: The quantity of the card in the order.
-   `name`: The name of the card.
-   `set`: The set code of the card.
-   `number`: The collector number of the card.
-   `condition`: The condition of the card.
-   `finish`: The finish of the card (e.g., foil, non-foil).
-   `price`: The price of the individual card.
-   `tcgplayer_sku`: The TCGPlayer SKU for the product, if available.
-   `scryfall_image_uri`: A URL to the card's image on Scryfall.
