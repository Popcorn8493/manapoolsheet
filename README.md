# Manapoolsheet

### Installation

Python 3.7+ https://www.python.org/

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
