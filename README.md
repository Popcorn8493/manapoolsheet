# Simple Manapoolsheet

Process Magic: The Gathering card inventory from ShipStation CSV files. Generates organized reports with card images and location mapping.

## Features

- Fetches card images from Scryfall API
- Maps card sets to physical storage locations
- Multi-level sorting (primary, secondary, tertiary)
- Generates CSV and interactive HTML reports
- Interactive HTML with collapsible sections and checkboxes
- Auto-opens HTML reports in browser (configurable)

<img width="1050" height="550" alt="image" src="https://github.com/user-attachments/assets/08937009-06ae-42bc-a87e-5bc2d4a8066a" />


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

1. **ShipStation CSV**: Tool automatically searches your Downloads folder for files matching:

   - `shipstation_orders*.csv`
   - `shipstation*.csv`
   - `*shipstation*.csv`

   Uses the most recent file found.

2. **Drawer Inventory JSON** (`inventory_locations.json`):

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
5. **Browser Configuration** (`config.json`): Auto-created with default settings

### Setup Inventory Mapping

Copy the template and edit with your set codes and drawer locations:

```bash
cp templates/inventory_locations_template.json inventory_locations.json
```

## Browser Configuration

The tool automatically opens HTML reports in your browser. Configure via `config.json` (auto-created) or command line:

**Config file** (`config.json`):
```json
{
  "browser": {
    "default_browser": "edge",
    "auto_open": true
  }
}
```

**Supported browsers**: `edge`, `chrome`, `firefox`, `default`

**Command line override**:
```bash
python manapoolsheet.py --browser chrome --no-open
```

## Usage

### Basic

```bash
python manapoolsheet.py
```

### Command Line Options

| Flag                | Description                    | Choices                                                                              | Default     |
| ------------------- | ------------------------------ | ------------------------------------------------------------------------------------ | ----------- |
| `--sort-by`         | Primary sort field             | `location`, `set`, `name`, `condition`, `rarity`, `price`, `card_type`, `color`      | `location`  |
| `--order`           | Primary sort order             | `asc`, `desc`                                                                        | `desc`      |
| `--secondary-sort`  | Secondary sort field           | Same as `--sort-by`                                                                  | None        |
| `--secondary-order` | Secondary sort order           | `asc`, `desc`                                                                        | `asc`       |
| `--tertiary-sort`   | Tertiary sort field            | Same as `--sort-by`                                                                  | None        |
| `--tertiary-order`  | Tertiary sort order            | `asc`, `desc`                                                                        | `asc`       |
| `--browser`         | Browser for HTML report        | `edge`, `chrome`, `firefox`, `default`                                               | From config |
| `--no-open`         | Disable auto-opening           | -                                                                                    | -           |
| `--clean-reports`   | Delete old reports             | -                                                                                    | -           |
| `--lionseye-export` | Generate Lion's Eye CSV export | -                                                                                    | -           |

### Examples

```bash
# Basic usage (sorts by location descending, opens in Edge)
python manapoolsheet.py

# Sort by name, ascending
python manapoolsheet.py --sort-by name --order asc

# Sort by price, descending
python manapoolsheet.py --sort-by price --order desc

# Multi-level sort: color → card type → name
python manapoolsheet.py --sort-by color --secondary-sort card_type --tertiary-sort name --order asc

# Sort by set, then by name
python manapoolsheet.py --sort-by set --secondary-sort name

# Generate Lion's Eye export for inventory updates
python manapoolsheet.py --lionseye-export

# Open in Chrome
python manapoolsheet.py --browser chrome

# Disable auto-opening
python manapoolsheet.py --no-open

# Clean up old reports (keeps most recent)
python manapoolsheet.py --clean-reports
```


### Output Files

- **CSV Report**: `csv_reports/card_inventory_report_YYYYMMDD_HHMMSS.csv`
- **HTML Report**: `html_reports/manapoolshoot_YYYYMMDD_HHMMSS.html`
- **Lion's Eye Export**: `lionseye_exports/lionseye_export_YYYYMMDD_HHMMSS.csv` (with `--lionseye-export`)
- **Card Images**: `card_images/` directory (cached for reuse)

## License

Personal use only. Respect Scryfall's API terms of service.
