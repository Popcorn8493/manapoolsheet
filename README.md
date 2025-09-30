# Simple Manapoolsheet

Process Magic: The Gathering card inventory from ShipStation CSV files. Generates organized reports with card images and location mapping.

## Features

- Fetches card images from Scryfall API
- Maps card sets to physical storage locations
- Sort by location, set, name, condition, rarity, or price
- Generates CSV and interactive HTML reports
- Interactive HTML with collapsible sections and checkboxes

## Setup

### Prerequisites
- Python 3.7+
- Virtual environment (recommended)

### Installation

1. **Create and activate virtual environment**
   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # Windows
   # source .venv/bin/activate  # macOS/Linux
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

### Required Files

1. **ShipStation CSV**: Place in project directory. Tool finds files matching:
   - `shipstation_orders*.csv`
   - `shipstation*.csv` 
   - `*shipstation*.csv`

2. **Drawer Inventory JSON** (`drawer_inventory.json`):
   ```json
   {
     "drawer_mapping": {
       "LTR": "Drawer A",
       "IKO": "Drawer B",
       "MB2": "Drawer C"
     }
   }
   ```

3. **HTML Template** (`templates/html_template_manapoolsheet.html`): Included with the project
4. **Inventory Locations Template** (`templates/inventory_locations_template.json`): Template for drawer mapping

### Using the Inventory Template

Copy the template to create your inventory mapping:
```bash
cp templates/inventory_locations_template.json inventory_locations.json
```

Then edit `inventory_locations.json` with your actual set codes and drawer locations.

## Usage

### Basic
```bash
python manapoolsheet.py
```

### Sorting Options

| Flag | Description | Choices | Default |
|------|-------------|---------|---------|
| `--sort-by` | Primary sort field | `location`, `set`, `name`, `condition`, `rarity`, `price` | `location` |
| `--order` | Sort order | `asc`, `desc` | `asc` |
| `--secondary-sort` | Secondary sort field | Same as `--sort-by` | None |
| `--secondary-order` | Secondary sort order | `asc`, `desc` | `asc` |
| `--clean-reports` | Delete old reports | - | - |

### Examples
```bash
# Sort by name, ascending
python manapoolsheet.py --sort-by name --order asc

# Sort by price, descending
python manapoolsheet.py --sort-by price --order desc

# Sort by set, then by name
python manapoolsheet.py --sort-by set --secondary-sort name

# Clean up old reports
python manapoolsheet.py --clean-reports
```

### Output Files
- **CSV**: `card_inventory_report_YYYYMMDD_HHMMSS.csv`
- **HTML**: `manapoolshoot_YYYYMMDD_HHMMSS.html`
- **Images**: `card_images/` directory

## License

Personal use only. Respect Scryfall's API terms of service.
