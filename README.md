# ManaPool Order Retriever

A Python tool for sellers to automate the creation of a fulfillment list from their ManaPool account.

This script connects to the ManaPool and Scryfall APIs to retrieve orders, collect detailed data for each item, and enrich it with card images and physical storage locations. The primary goal is to create a consolidated CSV file that can be used for efficient picking and shipping of orders.

## Features

-   **Order Management**: Fetches all or filtered (shipped/unshipped) orders from the ManaPool API.
-   **Detailed Card Information**: Retrieves comprehensive card data for each item in orders.
-   **Image Support**: Two options for card images:
    -   **URLs Only**: Default behavior - includes Scryfall image URLs in CSV (fast, minimal storage)
    -   **Local Downloads**: Optional feature - downloads card images locally for offline access
-   **Smart Location Mapping**: Maps card sets to user-defined physical locations with interactive setup.
-   **High-Value Card Alerts**: Identifies and highlights cards worth $10+ that may need special handling.
-   **Secure Credential Management**: Uses `.env` file for API credentials.
-   **Comprehensive Logging**: Logs all operations to `order_retrieval_log.txt`.
-   **Intelligent Caching**: Avoids re-downloading images that already exist locally.

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
    If a set from an order is not found in this file, the script will prompt you to assign a location interactively. You can either select from existing locations or create a new one.

### 3. Image Download Configuration (Optional)

You can configure image downloading behavior in two ways:

#### Option 1: Environment Variable
Add to your `.env` file:
```
DOWNLOAD_IMAGES=true   # Enable automatic image downloads
DOWNLOAD_IMAGES=false  # Disable image downloads (default)
```

#### Option 2: Interactive Prompt
If not set in the environment, the script will prompt you each time it runs whether to download images locally.

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

### First Run Setup
4.  On your first run, you will be prompted to configure:
    -   **Image Downloads**: Whether to download card images locally (Y/N)
    -   **HTML Reports**: Whether to generate interactive HTML reports (Y/N)
5.  These preferences are saved automatically for future runs.

### Subsequent Runs
The script will use your saved preferences automatically. To change them:
-   Type `reset` instead of selecting 1-3 when choosing which orders to retrieve
-   Or delete the `settings.json` file to start fresh

6.  The script will then fetch the order data, process it, and create output files in the project directory.

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
-   `scryfall_image_uri`: Either a URL to the card's image on Scryfall, or a local file path if images were downloaded.

## HTML Visual Reports

When enabled, the script generates an interactive HTML report alongside the CSV file. This provides:

### Features
-   **Visual Card Display**: Shows card images in an easy-to-browse format
-   **Location-Based Organization**: Cards grouped by physical storage location
-   **Progress Tracking**: Interactive checkboxes to mark items as picked
-   **Dual View Modes**:
    -   **Location View**: Default view organized by storage locations (optimized for picking)
    -   **Order View**: Switch to view cards grouped by individual orders (optimized for packing/shipping)
-   **Quick Navigation**: Switch between views with one click
-   **Order Integration**: Direct links to ManaPool seller order pages
-   **High-Value Alerts**: Special highlighting for cards worth $10+

### Usage
1.  Open the generated HTML file in your web browser
2.  Use checkboxes to track picking progress as you collect inventory
3.  Click "📦 View by Orders" to switch to order-grouped view for packing
4.  Click "← Back to Location View" to return to location-based organization
5.  Click "✅ Finished - Close & Delete" when complete

### View Modes Explained
-   **Location View**: Start here for efficient picking by walking through your storage areas
-   **Order View**: Switch here after picking to efficiently pack orders by grouping cards that go together

## Image Download Feature

The script supports two modes for card images:

### URL Mode (Default)
- Fast execution
- Minimal local storage
- Requires internet access to view images
- Perfect for users who don't need offline access

### Local Download Mode
- **First Run**: Slower as images are downloaded (with intelligent caching)
- **Subsequent Runs**: Fast as existing images are reused
- Works offline once images are downloaded
- Better for users who need reliable offline access
- Images stored in `images/` directory with consistent naming

### Image Caching System
When downloading images locally, the script:
- Creates unique identifiers for each card (set + collector number + name)
- Stores a cache mapping in `image_cache.json`
- Skips downloading images that already exist
- Verifies existing files and re-downloads if corrupted/missing
- Updates CSV with local file paths instead of URLs

### Storage Requirements
- Each card image: ~100-200KB
- For 1000 unique cards: ~100-200MB storage
- Cache file: <1MB typically

## File Structure

After running the script with image downloads enabled, your directory will contain:

```
├── manapoolsheet.py           # Main script
├── .env                       # Your credentials (create this)
├── locations.json             # Set-to-location mapping (create this)
├── image_cache.json           # Image download cache (auto-created)
├── images/                    # Downloaded card images (auto-created)
│   ├── M21_123_Lightning_Bolt.jpg
│   ├── ZNR_456_Island.jpg
│   └── ...
├── fulfillment_list_*.csv     # Generated order lists
└── order_retrieval_log.txt    # Application logs
```

## Advanced Configuration

### Environment Variables
Your `.env` file can contain:
```bash
# Required
MANAPOOL_EMAIL="your_email@example.com"
MANAPOOL_API_KEY="your_api_key"

# Optional
DOWNLOAD_IMAGES=true    # or false (default)
```

### Troubleshooting

**Image Download Issues:**
- Check internet connection for Scryfall API access
- Verify sufficient disk space for image storage
- Check `order_retrieval_log.txt` for specific error messages

**Performance Optimization:**
- First run with images enabled will be slower
- Subsequent runs reuse cached images for better performance
- Consider running with images disabled for quick order checks

**Storage Management:**
- Images are stored permanently until manually deleted
- Safe to delete `images/` folder - will re-download as needed
- `image_cache.json` tracks which images exist locally
