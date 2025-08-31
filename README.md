# Manapoolsheet

A Python tool for ManaPool sellers to automate the creation of fulfillment lists from their ManaPool account.

This script connects to the ManaPool and Scryfall APIs to retrieve orders, collect detailed data for each item, and enrich it with card images and physical storage locations. The primary goal is to create a consolidated CSV file that can be used for efficient picking and shipping of orders.

## Quick Start

### Prerequisites

* Python 3.x

### Installation

1.  Clone the repo
    ```sh
    git clone [https://github.com/Popcorn8493/manapoolsheet.git](https://github.com/Popcorn8493/manapoolsheet.git)
    ```
2.  Create a virtual environment
    ```sh
    python3 -m venv venv
    ```
3.  Activate the virtual environment
    * **Windows:**
        ```sh
        .\venv\Scripts\activate
        ```
    * **macOS/Linux:**
        ```sh
        source venv/bin/activate
        ```
4.  Install the required packages
    ```sh
    pip install -r requirements.txt
    ```


### Configuration

1. **API Credentials**: Create a `.env` file in the project root, and replace the example with your personal API details
   ```
   MANAPOOL_EMAIL="your_email@example.com"
   MANAPOOL_API_KEY="your_manapool_api_key"
   ```
   Get your API key and API email from: https://manapool.com/seller/integrations/manapool-api

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
