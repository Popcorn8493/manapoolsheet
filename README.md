# Manapoolsheet

### Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/Popcorn8493/manapoolsheet.git
   cd manapoolsheet
   ```

2. **Create virtual environment**

   ```bash
   python -m venv .venv
   ```

3. **Activate virtual environment**
   - **Windows:** `.\.venv\Scripts\activate`
   - **macOS/Linux:** `source .venv/bin/activate`

4. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

### Usage

1. **Run the application**

   ```bash
   python picklist_gui.py
   ```

2. **Load your picklist**
   - Click "Load CSV" to import your ManaPool picklist
   - The app automatically fetches card images and prices

3. **Manage your cards**
   - View cards in a responsive grid layout
   - Sort by name, location, price, or other criteria
   - Mark cards as "grabbed" when collected
   - Use the Summary tab for progress tracking

## File Formats

### Picklist CSV

The application expects CSV files with these columns:
- Order, Card Name, Set, Set Code, Collector #, Quantity, Condition, Language, Finish, Rarity

### Locations Configuration

Create a `locations.json` file to map set codes to physical storage locations:

```json
{
  "LTC": "Commander Binder A",
  "LTR": "Main Set Binder B", 
  "J22": "Jumpstart Binder C"
}
```

## Advanced Features

- **Responsive Grid Layout**: Cards automatically arrange based on window size
- **Price Integration**: Real-time market prices from Scryfall API
- **Smart Sorting**: Sort by price, location, rarity, or custom criteria
- **Progress Persistence**: Save and load your collection progress
- **Modern UI**: Clean, intuitive interface with visual feedback

## Troubleshooting

- **Images not loading**: Check your internet connection
- **Prices not showing**: Ensure you have a stable internet connection for API calls
- **CSV format issues**: Verify your CSV has the required column headers