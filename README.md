# Manapoolsheet

A Python tool for ManaPool sellers to automate the creation of fulfillment lists from their ManaPool account.

This script connects to the ManaPool and Scryfall APIs to retrieve orders, collect detailed data for each item, and enrich it with card images and physical storage locations. The primary goal is to create a consolidated CSV file that can be used for efficient picking and shipping of orders.

## Features

- **Order Management**: Fetches all or filtered (shipped/unshipped) orders from the ManaPool API
- **Detailed Card Information**: Retrieves comprehensive card data for each item in orders
- **Interactive HTML Reports**: Visual reports with dual view modes for picking and packing
- **Smart Location Mapping**: Maps card sets to user-defined physical locations with interactive setup
- **Image Support**: Optional local image downloads with intelligent caching
- **High-Value Card Alerts**: Identifies and highlights cards worth $10+ that may need special handling
- **Progress Tracking**: Interactive checkboxes to track fulfillment progress
- **User Preferences**: Saves settings for streamlined repeated use
- **Secure Credential Management**: Uses .env file for API credentials
- **Comprehensive Logging**: Logs all operations for troubleshooting

## Quick Start

### Prerequisites
- Python 3.6 or higher
- ManaPool seller account with API access

### Installation

1. Clone or download this repository
   ```bash
   git clone https://github.com/Popcorn8493/manapoolsheet.git
   ```
   
3. Navigate to the project directory in your terminal
    ```bash
   cd manapoolsheet
   ```
5. Install required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Configuration

1. **API Credentials**: Create a `.env` file in the project root:
   ```
   MANAPOOL_EMAIL="your_email@example.com"
   MANAPOOL_API_KEY="your_manapool_api_key"
   ```
   Get your API key from: https://manapool.com/seller/integrations/manapool-api

2. **Location Mapping** (Optional): The script will automatically create a `locations.json` file as you assign locations to new sets. You can also pre-create this file if desired:
   ```json
   {
     "M21": "Shelf A",
     "ZNR": "Shelf B", 
     "KHM": "Box 1",
     "STX": "Binder 1"
   }
   ```

### Usage

Run the script:
```bash
python manapoolsheet.py
```

**First Run:**
- Choose which orders to retrieve (Not Shipped/Shipped/All)
- Configure image downloads (Y/N)
- Configure HTML reports (Y/N)
- Settings are saved for future runs

**Subsequent Runs:**
- Just select which orders to retrieve
- Type 'reset' to change saved preferences

## Output Files

### CSV Export
Contains detailed order data with columns:
- Order ID and label
- Physical location (based on set mapping)
- Card details (name, set, number, condition, finish)
- Pricing and quantity information
- Scryfall image URLs
- TCGPlayer SKU

### HTML Visual Report (Optional)
Interactive web-based report featuring:
- **Location View**: Cards organized by storage location for efficient picking
- **Order View**: Cards grouped by order for efficient packing
- Progress tracking with checkboxes
- Card images and detailed information
- Direct links to ManaPool order pages
- High-value card highlighting

## Advanced Features

### Smart Location Assignment
- Automatically prompts for new sets not in your location mapping
- Provides autocomplete suggestions from existing locations
- Updates location file automatically

### Image Management
- **URL Mode**: Fast, minimal storage (default)
- **Local Download Mode**: Offline access, cached for performance
- Intelligent caching prevents re-downloading existing images
- Images stored in organized directory structure

### User Preferences
- Automatically saves image download and HTML report preferences
- Reset option available when needed
- Environment variable override support

## File Structure

After setup, your directory will contain:
```
├── manapoolsheet.py           # Main script
├── requirements.txt           # Python dependencies
├── .env                      # Your credentials (create this)
├── locations.json            # Set-to-location mapping (optional)
├── settings.json             # User preferences (auto-created)
├── fulfillment_list_*.csv    # Generated order lists
├── fulfillment_list_*.html   # Visual reports (if enabled)
└── order_retrieval_log.txt   # Application logs
```

## Sorting and Organization

The CSV output is intelligently sorted by:
1. Location with most total quantity first
2. Alphabetical location names (for consistency)
3. Sets with most quantity within each location first
4. Alphabetical set names (for consistency)  
5. Alphabetical card names

This organization optimizes the picking workflow by prioritizing high-volume areas and maintaining consistent ordering.

## Troubleshooting

**Common Issues:**
- **API Errors**: Verify credentials in .env file
- **Missing Images**: Check internet connection for Scryfall API access
- **Performance**: First run with images enabled will be slower
- **Storage**: Each image is ~100-200KB, plan accordingly

**Getting Help:**
- Check `order_retrieval_log.txt` for detailed error messages
- Ensure Python 3.6+ is installed
- Verify all required packages are installed via requirements.txt

## Environment Variables

Your `.env` file supports:
```bash
# Required
MANAPOOL_EMAIL="your_email@example.com"
MANAPOOL_API_KEY="your_api_key"

# Optional
DOWNLOAD_IMAGES=true    # Force enable/disable image downloads
```

## Performance Tips

- **Quick Checks**: Disable images and HTML for fast order verification
- **First Setup**: Enable images on first run, then reuse cached images
- **Storage Management**: Delete `images/` folder to reclaim space (will re-download as needed)
- **Large Inventories**: Consider using location mapping to organize efficiently

## Security Notes

- API credentials are stored locally in .env file
- .env file is excluded from version control via .gitignore
- No credentials are transmitted except to official ManaPool/Scryfall APIs
- Generated files contain only order data you already have access to

## License

This project is open source. See the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.
