# Simple Manapoolsheet

Process Magic: The Gathering card inventory from ShipStation CSV files. Generates organized reports with card images and location mapping.

## Features

- Fetches card images from Scryfall API
- Maps card sets to physical storage locations
- Sort by location, set, name, condition, rarity, or price
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

1. **ShipStation CSV**: Place in project directory. Tool finds files matching:

   - `shipstation_orders*.csv`
   - `shipstation*.csv`
   - `*shipstation*.csv`

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

### Using the Inventory Template

Copy the template to create your inventory mapping:

```bash
cp templates/inventory_locations_template.json inventory_locations.json
```

Then edit `inventory_locations.json` with your actual set codes and drawer locations.

## Browser Configuration

The tool automatically opens HTML reports in your browser. You can configure this behavior:

### Default Settings

- **Default Browser**: Microsoft Edge
- **Auto-Open**: Enabled
- **Config File**: `config.json` (auto-created)

### Configuration Options

Edit `config.json` to change browser settings:

```json
{
  "browser": {
    "default_browser": "edge",
    "auto_open": true
  }
}
```

**Supported Browsers**: `edge`, `chrome`, `firefox`, `default`

### Command Line Override

Override config settings via command line:

```bash
# Use specific browser
python manapoolsheet.py --browser chrome

# Disable auto-opening
python manapoolsheet.py --no-open

# Use Firefox and disable auto-open
python manapoolsheet.py --browser firefox --no-open
```

## Usage

### Basic

```bash
python manapoolsheet.py
```

### Command Line Options

| Flag                | Description             | Choices                                                   | Default     |
| ------------------- | ----------------------- | --------------------------------------------------------- | ----------- |
| `--sort-by`         | Primary sort field      | `location`, `set`, `name`, `condition`, `rarity`, `price` | `location`  |
| `--order`           | Sort order              | `asc`, `desc`                                             | `asc`       |
| `--secondary-sort`  | Secondary sort field    | Same as `--sort-by`                                       | None        |
| `--secondary-order` | Secondary sort order    | `asc`, `desc`                                             | `asc`       |
| `--browser`         | Browser for HTML report | `edge`, `chrome`, `firefox`, `default`                    | From config |
| `--no-open`         | Disable auto-opening    | -                                                         | -           |
| `--clean-reports`   | Delete old reports      | -                                                         | -           |

### Examples

```bash
# Basic usage (opens in Edge by default)
python manapoolsheet.py

# Sort by name, ascending
python manapoolsheet.py --sort-by name --order asc

# Sort by price, descending
python manapoolsheet.py --sort-by price --order desc

# Sort by color, type, name

python manapoolsheet.py --sort-by color --secondary-sort card_type --tertiary-sort name --order asc

# Sort by set, then by name
python manapoolsheet.py --sort-by set --secondary-sort name

# Open in Chrome
python manapoolsheet.py --browser chrome

# Disable auto-opening
python manapoolsheet.py --no-open

# Clean up old reports
python manapoolsheet.py --clean-reports
```


### Output Files

- **CSV**: `card_inventory_report_YYYYMMDD_HHMMSS.csv`
- **HTML**: `manapoolshoot_YYYYMMDD_HHMMSS.html`
- **Images**: `card_images/` directory

## License

Personal use only. Respect Scryfall's API terms of service.
