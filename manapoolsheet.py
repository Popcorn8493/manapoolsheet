import csv
import json
import requests
import time
import os
import glob
import argparse
import webbrowser
import subprocess
import platform
from urllib.parse import urlparse
from typing import Dict, List, Optional, Callable, Any
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

CACHE_DIR = "card_images"
DRAWER_FILE = "inventory_locations.json"
TEMPLATE_FILE = "templates/html_template_manapoolsheet.html"
CSV_OUTPUT_DIR = "csv_reports"
HTML_OUTPUT_DIR = "html_reports"
CONFIG_FILE = "config.json"
CSV_FIELDS = {
    "card_name": "Card Name",
    "set_code": "Set Code",
    "collector_number": "Collector #",
    "set_name": "Set",
    "quantity": "Quantity",
    "condition": "Condition",
    "rarity": "Rarity",
    "finish": "Finish",
    "unit_price": "Unit Price",
}
SHIPSTATION_PATTERNS = [
    "shipstation_orders*.csv",
    "shipstation*.csv",
    "*shipstation*.csv",
]
REPORT_PATTERNS = {"csv": "card_inventory_report_*.csv", "html": "manapoolshoot_*.html"}
SCRYFALL_BASE_URL = "https://api.scryfall.com/cards/search"
REQUEST_TIMEOUT = 30
SCRYFALL_TIMEOUT = 10
IMAGE_DOWNLOAD_DELAY = 0.1
_session = None


def get_session() -> requests.Session:
    global _session
    if _session is None:
        _session = requests.Session()
    return _session


@dataclass
class Card:
    name: str
    set_name: str
    set_code: str
    collector_number: str
    quantity: str
    condition: str
    rarity: str
    finish: str
    location: str
    image_url: str
    price: str


def get_downloads_folder() -> str:
    home = Path.home()
    downloads_paths = [
        home / "Downloads",
        home / "downloads",
        Path(os.path.expanduser("~/Downloads")),
        Path(os.path.expanduser("~/downloads")),
    ]
    if os.name == "nt":
        downloads_paths.extend(
            [
                Path(os.path.expanduser("~/Downloads")),
                Path(os.path.expanduser("~/downloads")),
            ]
        )
    for path in downloads_paths:
        if path.exists() and path.is_dir():
            return str(path)
    return str(home)


def find_most_recent_shipstation_file() -> Optional[str]:
    downloads_folder = get_downloads_folder()
    print(f"Searching for ShipStation files in: {downloads_folder}")
    all_files = []
    for pattern in SHIPSTATION_PATTERNS:
        search_path = os.path.join(downloads_folder, pattern)
        all_files.extend(glob.glob(search_path))
    if not all_files:
        print("No ShipStation files found in Downloads folder")
        return None
    most_recent = max(all_files, key=os.path.getmtime)
    print(
        f"Found {len(all_files)} ShipStation files, using most recent: {os.path.basename(most_recent)}"
    )
    return most_recent


def parse_price(price_str: str) -> float:
    try:
        if price_str and price_str != "N/A":
            return float(price_str.replace("$", "").replace(",", ""))
        return 0.0
    except (ValueError, AttributeError):
        return 0.0


def parse_quantity(quantity_str: str) -> int:
    try:
        return int(quantity_str)
    except (ValueError, TypeError):
        return 1


def sanitize_for_filename(text: str) -> str:
    safe_text = "".join(
        (c for c in text if c.isalnum() or c in (" ", "-", "_"))
    ).rstrip()
    return safe_text.replace(" ", "_")


def infer_file_extension(image_url: str) -> str:
    parsed_url = urlparse(image_url)
    return os.path.splitext(parsed_url.path)[1] or ".jpg"


def build_scryfall_queries(
    card_name: str, set_code: str, collector_number: Optional[str] = None
) -> List[str]:
    clean_name = (
        card_name.replace('"', "").replace("\n", " ").replace("\r", " ").strip()
    )
    if collector_number:
        return [
            f"cn:{collector_number} set:{set_code}",
            f'name:"{clean_name}" cn:{collector_number} set:{set_code}',
            f'name:"{clean_name}" set:{set_code}',
        ]
    else:
        return [
            f'name:"{clean_name}" set:{set_code}',
            f"name:{clean_name} set:{set_code}",
        ]


def create_sort_key_factory(sort_field: str) -> Callable[[Card], Any]:

    def get_sort_key(card: Card):
        if sort_field == "location":
            return card.location
        elif sort_field == "set":
            return card.set_name
        elif sort_field == "name":
            return card.name.lower()
        elif sort_field == "condition":
            return card.condition
        elif sort_field == "rarity":
            return card.rarity
        elif sort_field == "price":
            return parse_price(card.price)
        return ""

    return get_sort_key


def load_drawer_mapping(json_file: str) -> Dict[str, str]:
    with open(json_file, "r", encoding="utf-8") as file:
        data = json.load(file)
        return data.get("drawer_mapping", {})


def load_drawer_mapping_safe(json_file: str) -> Dict[str, str]:
    try:
        return load_drawer_mapping(json_file)
    except FileNotFoundError:
        return {}


def load_browser_config(config_file: str = CONFIG_FILE) -> Dict[str, Any]:
    try:
        with open(config_file, "r", encoding="utf-8") as file:
            data = json.load(file)
            return data.get("browser", {})
    except FileNotFoundError:
        print(
            f"Warning: Config file {config_file} not found. Using default browser settings."
        )
        return {"default_browser": "edge", "auto_open": True, "browser_options": {}}
    except json.JSONDecodeError as e:
        print(
            f"Warning: Error parsing config file {config_file}: {e}. Using default settings."
        )
        return {"default_browser": "edge", "auto_open": True, "browser_options": {}}


def open_html_in_browser(
    html_file: str, browser: str = "edge", auto_open: bool = True
) -> bool:
    if not auto_open:
        print(f"Auto-open disabled. HTML report available at: {html_file}")
        return False
    if not os.path.exists(html_file):
        print(f"Error: HTML file {html_file} not found!")
        return False
    try:
        abs_path = os.path.abspath(html_file)
        file_url = f"file:///{abs_path.replace(os.sep, '/')}"
        if browser.lower() == "edge":
            try:
                if platform.system() == "Windows":
                    subprocess.run(
                        ["start", "msedge", file_url], check=True, shell=True
                    )
                else:
                    subprocess.run(["msedge", file_url], check=True)
                print(f"Opened HTML report in Microsoft Edge: {html_file}")
                return True
            except (subprocess.CalledProcessError, FileNotFoundError):
                print("Microsoft Edge not found, falling back to default browser...")
                webbrowser.open(file_url)
                print(f"Opened HTML report in default browser: {html_file}")
                return True
        elif browser.lower() == "chrome":
            try:
                if platform.system() == "Windows":
                    subprocess.run(
                        ["start", "chrome", file_url], check=True, shell=True
                    )
                else:
                    subprocess.run(["google-chrome", file_url], check=True)
                print(f"Opened HTML report in Chrome: {html_file}")
                return True
            except (subprocess.CalledProcessError, FileNotFoundError):
                print("Chrome not found, falling back to default browser...")
                webbrowser.open(file_url)
                print(f"Opened HTML report in default browser: {html_file}")
                return True
        elif browser.lower() == "firefox":
            try:
                if platform.system() == "Windows":
                    subprocess.run(
                        ["start", "firefox", file_url], check=True, shell=True
                    )
                else:
                    subprocess.run(["firefox", file_url], check=True)
                print(f"Opened HTML report in Firefox: {html_file}")
                return True
            except (subprocess.CalledProcessError, FileNotFoundError):
                print("Firefox not found, falling back to default browser...")
                webbrowser.open(file_url)
                print(f"Opened HTML report in default browser: {html_file}")
                return True
        else:
            webbrowser.open(file_url)
            print(f"Opened HTML report in default browser: {html_file}")
            return True
    except Exception as e:
        print(f"Error opening browser: {e}")
        print(f"HTML report available at: {html_file}")
        return False


def parse_arguments():
    parser = argparse.ArgumentParser(description="Generate card inventory reports")
    parser.add_argument(
        "--sort-by",
        choices=[
            "location",
            "set",
            "name",
            "condition",
            "rarity",
            "price",
            "card_type",
            "color",
        ],
        default="location",
        help="Field to sort by (default: location)",
    )
    parser.add_argument(
        "--order",
        choices=["asc", "desc"],
        default="desc",
        help="Sort order: asc for ascending, desc for descending (default: desc)",
    )
    parser.add_argument(
        "--secondary-sort",
        choices=[
            "location",
            "set",
            "name",
            "condition",
            "rarity",
            "price",
            "card_type",
            "color",
        ],
        help="Secondary sort field (optional)",
    )
    parser.add_argument(
        "--secondary-order",
        choices=["asc", "desc"],
        default="asc",
        help="Secondary sort order (default: asc)",
    )
    parser.add_argument(
        "--tertiary-sort",
        choices=[
            "location",
            "set",
            "name",
            "condition",
            "rarity",
            "price",
            "card_type",
            "color",
        ],
        help="Tertiary sort field (optional)",
    )
    parser.add_argument(
        "--tertiary-order",
        choices=["asc", "desc"],
        default="asc",
        help="Tertiary sort order (default: asc)",
    )
    parser.add_argument(
        "--clean-reports",
        action="store_true",
        help="Delete all but the most recent CSV and HTML reports",
    )
    parser.add_argument(
        "--browser",
        choices=["edge", "chrome", "firefox", "default"],
        help="Browser to open HTML report in (overrides config file)",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Do not automatically open HTML report in browser",
    )
    return parser.parse_args()


def sort_cards(
    cards: List[Card],
    sort_by: str,
    order: str,
    secondary_sort: Optional[str] = None,
    secondary_order: str = "asc",
) -> List[Card]:
    from collections import Counter

    get_sort_key = create_sort_key_factory(sort_by)
    get_secondary_key = (
        create_sort_key_factory(secondary_sort) if secondary_sort else None
    )
    if sort_by == "location":
        location_counts = Counter((card.location for card in cards))

        def get_location_count(card):
            return location_counts[card.location]

        reverse = order == "desc"
        if secondary_sort and get_secondary_key:
            cards.sort(
                key=lambda x: (get_location_count(x), get_secondary_key(x)),
                reverse=reverse,
            )
        else:
            cards.sort(key=get_location_count, reverse=reverse)
    else:
        reverse = order == "desc"
        if secondary_sort and get_secondary_key:
            cards.sort(
                key=lambda x: (get_sort_key(x), get_secondary_key(x)), reverse=reverse
            )
        else:
            cards.sort(key=get_sort_key, reverse=reverse)
    return cards


def download_and_cache_image(
    image_url: str, card_name: str, set_code: str
) -> Optional[str]:
    if not image_url:
        return None
    os.makedirs(CACHE_DIR, exist_ok=True)
    safe_name = sanitize_for_filename(card_name)
    safe_set = set_code.replace("/", "_")
    file_ext = infer_file_extension(image_url)
    filename = f"{safe_name}_{safe_set}{file_ext}"
    local_path = os.path.join(CACHE_DIR, filename)
    if os.path.exists(local_path):
        print(f"Using cached image for {card_name} ({set_code})")
        return local_path
    try:
        print(f"Downloading image for {card_name} ({set_code})...")
        session = get_session()
        response = session.get(image_url, timeout=SCRYFALL_TIMEOUT)
        response.raise_for_status()
        with open(local_path, "wb") as f:
            f.write(response.content)
        print(f"Cached image for {card_name} ({set_code})")
        return local_path
    except Exception as e:
        print(f"Error downloading image for {card_name} ({set_code}): {e}")
        return None


def get_card_image_url(
    card_name: str, set_code: str, collector_number: Optional[str] = None
) -> Optional[str]:
    try:
        search_queries = build_scryfall_queries(card_name, set_code, collector_number)
        for query in search_queries:
            try:
                params = {"q": query, "format": "json"}
                session = get_session()
                response = session.get(
                    SCRYFALL_BASE_URL, params=params, timeout=SCRYFALL_TIMEOUT
                )
                response.raise_for_status()
                data = response.json()
                if data.get("data") and len(data["data"]) > 0:
                    card = data["data"][0]
                    image_uris = card.get("image_uris", {})
                    image_url = (
                        image_uris.get("large")
                        or image_uris.get("normal")
                        or image_uris.get("small")
                    )
                    if image_url:
                        return image_url
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 404:
                    continue
                else:
                    raise
            except Exception:
                continue
        return None
    except Exception as e:
        print(f"Error fetching image for {card_name} ({set_code}): {e}")
        return None


def process_shipstation_data(
    input_file: str, drawer_mapping: Dict[str, str]
) -> List[Card]:
    cards = []
    with open(input_file, "r", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if not row.get(CSV_FIELDS["card_name"]):
                continue
            set_code = row.get(CSV_FIELDS["set_code"], "")
            card_name = row.get(CSV_FIELDS["card_name"], "")
            collector_number = row.get(CSV_FIELDS["collector_number"], "")
            set_name = row.get(CSV_FIELDS["set_name"], "")
            quantity = row.get(CSV_FIELDS["quantity"], "1")
            condition = row.get(CSV_FIELDS["condition"], "")
            rarity = row.get(CSV_FIELDS["rarity"], "")
            finish = row.get(CSV_FIELDS["finish"], "")
            price = row.get(CSV_FIELDS["unit_price"], "")
            location = drawer_mapping.get(set_code, "Unknown")
            print(f"Fetching image for {card_name} ({set_code})...")
            image_url = get_card_image_url(card_name, set_code, collector_number)
            cached_image_path = (
                download_and_cache_image(image_url, card_name, set_code)
                if image_url
                else None
            )
            card = Card(
                name=card_name,
                set_name=set_name,
                set_code=set_code,
                collector_number=collector_number,
                quantity=quantity,
                condition=condition,
                rarity=rarity,
                finish=finish,
                location=location,
                image_url=cached_image_path or image_url or "",
                price=price,
            )
            cards.append(card)
            time.sleep(IMAGE_DOWNLOAD_DELAY)
    return cards


def cleanup_old_cache():
    if not os.path.exists(CACHE_DIR):
        return
    cached_files = set(os.listdir(CACHE_DIR))
    print(f"Cache directory contains {len(cached_files)} images")


def cleanup_old_reports():
    csv_pattern = os.path.join(CSV_OUTPUT_DIR, "card_inventory_report_*.csv")
    html_pattern = os.path.join(HTML_OUTPUT_DIR, "manapoolshoot_*.html")
    csv_files = glob.glob(csv_pattern)
    html_files = glob.glob(html_pattern)
    csv_files.sort(key=os.path.getmtime, reverse=True)
    html_files.sort(key=os.path.getmtime, reverse=True)
    files_to_delete = []
    if len(csv_files) > 1:
        files_to_delete.extend(csv_files[1:])
    if len(html_files) > 1:
        files_to_delete.extend(html_files[1:])
    deleted_count = 0
    for file_path in files_to_delete:
        try:
            os.remove(file_path)
            deleted_count += 1
            print(f"Deleted old report: {file_path}")
        except OSError as e:
            print(f"Error deleting {file_path}: {e}")
    if deleted_count > 0:
        print(f"Cleaned up {deleted_count} old report files")
    else:
        print("No old report files to clean up")


def generate_csv_report(cards: List[Card], output_file: str):
    with open(output_file, "w", newline="", encoding="utf-8") as file:
        fieldnames = [
            "Card Name",
            "Set",
            "Set Code",
            "Collector #",
            "Quantity",
            "Condition",
            "Rarity",
            "Location",
            "Unit Price",
            "Image URL",
        ]
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for card in cards:
            writer.writerow(
                {
                    "Card Name": card.name,
                    "Set": card.set_name,
                    "Set Code": card.set_code,
                    "Collector #": card.collector_number,
                    "Quantity": card.quantity,
                    "Condition": card.condition,
                    "Rarity": card.rarity,
                    "Location": card.location,
                    "Unit Price": card.price,
                    "Image URL": card.image_url,
                }
            )


def get_card_highlight_classes(card: Card) -> str:
    classes = []
    try:
        quantity = parse_quantity(card.quantity)
        if quantity > 1:
            classes.append("highlight-quantity")
    except (ValueError, TypeError):
        pass
    finish = card.finish.lower()
    if "foil" in finish and "non-foil" not in finish:
        classes.append("highlight-foil")
    if "etched" in finish:
        classes.append("highlight-etched")
    return " ".join(classes)


def load_html_template(template_file: str = TEMPLATE_FILE) -> Optional[str]:
    try:
        with open(template_file, "r", encoding="utf-8") as file:
            return file.read()
    except FileNotFoundError:
        print(f"Error: Template file {template_file} not found!")
        return None


def get_image_path_for_html(image_path: str) -> str:
    if not image_path:
        return ""
    if image_path.startswith(("http://", "https://")):
        return image_path
    if os.path.isabs(image_path):
        return f"file://{image_path.replace(os.sep, '/')}"
    else:
        if image_path.startswith(CACHE_DIR + os.sep) or image_path.startswith(
            CACHE_DIR + "/"
        ):
            filename = os.path.basename(image_path)
            return f"../{CACHE_DIR}/{filename}"
        else:
            return f"../{CACHE_DIR}/{image_path}"


def render_card_html(card: Card) -> str:
    if card.image_url:
        html_image_path = get_image_path_for_html(card.image_url)
        image_html = f'<div class="card-image-container"><img src="{html_image_path}" alt="{card.name}" class="card-image"></div>'
    else:
        image_html = '<div class="no-image">No image available</div>'
    highlight_classes = get_card_highlight_classes(card)
    card_classes = f"card-item {highlight_classes}".strip()
    safe_id = f"card_{card.name.replace(' ', '_')}_{card.set_code}"
    return f"""\n                <div class="{card_classes}">\n                    <input type="checkbox" class="card-checkbox"\n                           id="{safe_id}">\n                    <div class="found-indicator">FOUND</div>\n                    {image_html}\n                    <div class="card-info">\n                        <h4 class="card-name">{card.name}</h4>\n                        <div class="card-details">\n                            <strong>Set:</strong> {card.set_name}\n                            ({card.set_code})<br>\n                            <strong>Collector #:</strong>\n                            {card.collector_number}<br>\n                            <strong>Quantity:</strong> {card.quantity}<br>\n                            <strong>Condition:</strong> {card.condition}<br>\n                            <strong>Rarity:</strong> {card.rarity}<br>\n                            <strong>Finish:</strong> {card.finish}\n                        </div>\n                        <div class="pricing-info">\n                            <strong>Unit Price:</strong>\n                            ${(card.price if card.price else 'N/A')}\n                        </div>\n                    </div>\n                </div>\n"""


def render_location_section(location: str, cards: List[Card]) -> str:
    location_section = f'\n        <div class="location-section">\n            <h2 class="location-header">{location}</h2>\n            <div class="location-content">\n                <div class="cards-grid">\n'
    for card in cards:
        location_section += render_card_html(card)
    location_section += "\n                </div>\n            </div>\n        </div>\n"
    return location_section


def generate_html_report(
    cards: List[Card], output_file: str, sort_by: str = "location", order: str = "asc"
):
    from datetime import datetime

    template = load_html_template()
    if template is None:
        return
    if sort_by == "location":
        grouped_cards: Dict[str, List[Card]] = defaultdict(list)
        for card in cards:
            grouped_cards[card.location].append(card)
        location_order = []
        seen_locations = set()
        for card in cards:
            location = card.location
            if location not in seen_locations:
                location_order.append(location)
                seen_locations.add(location)
    else:
        grouped_cards = {"All Cards": cards}
        location_order = ["All Cards"]
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    location_sections_html = ""
    for location in location_order:
        cards_in_location = grouped_cards[location]
        location_sections_html += render_location_section(location, cards_in_location)
    total_cards = sum((parse_quantity(card.quantity) for card in cards))
    unique_cards = len(cards)
    unique_locations = len(set((card.location for card in cards)))
    locations_count = unique_locations
    html_content = template.replace("{timestamp}", timestamp)
    html_content = html_content.replace("{total_cards}", str(total_cards))
    html_content = html_content.replace("{unique_cards}", str(unique_cards))
    html_content = html_content.replace("{locations_count}", str(locations_count))
    html_content = html_content.replace("{location_sections}", location_sections_html)
    with open(output_file, "w", encoding="utf-8") as file:
        file.write(html_content)


def main():
    from datetime import datetime

    args = parse_arguments()
    browser_config = load_browser_config()
    browser = (
        args.browser if args.browser else browser_config.get("default_browser", "edge")
    )
    auto_open = not args.no_open and browser_config.get("auto_open", True)
    os.makedirs(CSV_OUTPUT_DIR, exist_ok=True)
    os.makedirs(HTML_OUTPUT_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_output_file = os.path.join(
        CSV_OUTPUT_DIR, f"card_inventory_report_{timestamp}.csv"
    )
    html_output_file = os.path.join(HTML_OUTPUT_DIR, f"manapoolshoot_{timestamp}.html")
    print("Looking for ShipStation files...")
    shipstation_file = find_most_recent_shipstation_file()
    if not shipstation_file:
        print("Error: No ShipStation CSV files found in Downloads folder!")
        print(
            "Looking for files matching: shipstation_orders*.csv, shipstation*.csv, *shipstation*.csv"
        )
        return
    print(f"Using ShipStation file: {shipstation_file}")
    print("Loading drawer mapping...")
    drawer_mapping = load_drawer_mapping_safe(DRAWER_FILE)
    if drawer_mapping:
        print(f"Loaded {len(drawer_mapping)} set-to-drawer mappings")
    else:
        print(f"Note: {DRAWER_FILE} not found. All cards will use 'Unknown' location.")
    print("Checking image cache...")
    cleanup_old_cache()
    print("Processing ShipStation data...")
    cards = process_shipstation_data(shipstation_file, drawer_mapping)
    print(f"Processed {len(cards)} cards")
    print(f"Sorting cards by {args.sort_by} ({args.order})...")
    if args.secondary_sort:
        print(f"Secondary sort by {args.secondary_sort} ({args.secondary_order})...")
    cards = sort_cards(
        cards, args.sort_by, args.order, args.secondary_sort, args.secondary_order
    )
    print("Generating CSV report...")
    generate_csv_report(cards, csv_output_file)
    print(f"CSV report generated: {csv_output_file}")
    print("Generating HTML report...")
    generate_html_report(cards, html_output_file, args.sort_by, args.order)
    print(f"HTML report generated: {html_output_file}")
    if args.clean_reports:
        print("\nCleaning up old reports...")
        cleanup_old_reports()
    print("\nReports generated successfully!")
    print(f"- CSV: {csv_output_file}")
    print(f"- HTML: {html_output_file}")
    print(f"- Images cached in: {CACHE_DIR}/")
    print(f"- CSV reports directory: {CSV_OUTPUT_DIR}/")
    print(f"- HTML reports directory: {HTML_OUTPUT_DIR}/")
    if auto_open:
        print(f"\nOpening HTML report in {browser}...")
        open_html_in_browser(html_output_file, browser, auto_open)
    else:
        print(f"\nHTML report available at: {html_output_file}")
        print(
            "Use --browser <browser> to specify a browser or remove --no-open to auto-open"
        )


if __name__ == "__main__":
    main()
