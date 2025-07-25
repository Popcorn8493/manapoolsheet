import sys
import os
import json
import requests
import logging
import pandas as pd
import time
import webbrowser
import html
from datetime import datetime
from typing import Any, Dict, List, Optional, Set
from dotenv import load_dotenv
import re

# --- Configuration ---

# Load environment variables from a .env file
load_dotenv()

LOG_FILE = "order_retrieval_log.txt"
BASE_OUTPUT_CSV_NAME = "fulfillment_list"
LOCATIONS_FILE = "locations.json"  # File to map set codes to physical locations
IMAGES_DIR = "images"  # Directory to store downloaded card images
IMAGE_CACHE_FILE = "image_cache.json"  # Cache file to track downloaded images
SETTINGS_FILE = "settings.json"  # User preferences file

# API and Header Configuration
MANAPOOL_API_BASE_URL = "https://manapool.com/api/v1"
SCRYFALL_API_BASE_URL = "https://api.scryfall.com"
FULFILLMENT_FIELD = "latest_fulfillment_status"
SHIPPED_VALUE = "shipped"

# --- Helper Functions ---


def load_set_locations(filename: str) -> Dict[str, str]:
    """Loads the set-to-location mapping from a JSON file."""
    try:
        with open(filename, "r", encoding="utf-8") as f:
            locations = json.load(f)
            # Ensure all keys (set codes) are uppercase for consistent matching
            return {k.upper(): v for k, v in locations.items()}
    except FileNotFoundError:
        logging.warning(
            f"Location file not found: {filename}. Location data will be missing."
        )
        return {}
    except json.JSONDecodeError:
        logging.error(
            f"Could not parse {filename}. Please ensure it is valid JSON.")
        return {}


def save_set_locations(filename: str, locations: Dict[str, str]) -> None:
    """Saves the set-to-location mapping to a JSON file."""
    try:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(locations, f, indent=2, sort_keys=True)
        logging.info(f"Updated locations saved to {filename}")
    except Exception as exc:
        logging.error(f"Failed to save locations to {filename}: {exc}")

def load_image_cache(filename: str) -> Dict[str, str]:
    """Load the image cache from a JSON file."""
    try:
        with open(filename, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_image_cache(filename: str, cache: Dict[str, str]) -> None:
    """Save the image cache to a JSON file."""
    try:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(cache, f, indent=2, sort_keys=True)
    except Exception as exc:
        logging.error(f"Failed to save image cache to {filename}: {exc}")


def load_user_settings(filename: str) -> Dict[str, Any]:
    """Load user settings from a JSON file."""
    try:
        with open(filename, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {
            "download_images": None,
            "generate_html_report": None
        }


def save_user_settings(filename: str, settings: Dict[str, Any]) -> None:
    """Save user settings to a JSON file."""
    try:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, sort_keys=True)
    except Exception as exc:
        logging.error(f"Failed to save user settings to {filename}: {exc}")


def generate_image_key(card_name: str, set_code: str, collector_number: str) -> str:
    """Generate a unique key for caching downloaded images."""
    # Create a consistent key that uniquely identifies a card
    return f"{set_code.upper()}_{collector_number}_{card_name.replace(' ', '_').replace('/', '_')}"


def download_card_image(image_url: str, image_key: str, images_dir: str) -> str:
    """Download an image and return the local file path."""
    if image_url == "N/A":
        return "N/A"
    
    # Create images directory if it doesn't exist
    os.makedirs(images_dir, exist_ok=True)
    
    # Generate filename with .jpg extension
    filename = f"{image_key}.jpg"
    filepath = os.path.join(images_dir, filename)
    
    # Check if file already exists
    if os.path.exists(filepath):
        return filepath
    
    try:
        # Download the image
        response = requests.get(image_url, timeout=30)
        response.raise_for_status()
        
        # Save the image
        with open(filepath, 'wb') as f:
            f.write(response.content)
        
        logging.info(f"Downloaded image: {filename}")
        return filepath
        
    except requests.exceptions.RequestException as exc:
        logging.warning(f"Failed to download image {image_url}: {exc}")
        return "N/A"
        

def natural_sort_key(s: str) -> List[Any]:
    """Key for natural sorting strings containing numbers."""
    return [int(text) if text.isdigit() else text.lower()
            for text in re.split(r'(\d+)', s)]

def get_location_for_set(set_code: str, existing_locations: Dict[str, str],
                        new_sets_found: Set[str]) -> str:
    """Get location for a set, prompting user for new sets with autocomplete."""
    set_upper = set_code.upper()
    
    # If we already have this set mapped, return it
    if set_upper in existing_locations:
        return existing_locations[set_upper]
    
    # If we've already asked about this set in this session, skip asking again
    if set_upper in new_sets_found:
        return existing_locations.get(set_upper, "Unassigned")
    
    # New set found - prompt user
    new_sets_found.add(set_upper)
    unique_locations = sorted(set(existing_locations.values()), key=natural_sort_key)
    
    print(f"\n📦 New set found: {set_code}")
    print("Existing locations:")
    for i, loc in enumerate(unique_locations, 1):
        print(f"  {i}: {loc}")
    print(f"  {len(unique_locations) + 1}: Create new location")
    
    while True:
        choice = input(f"Select location for {set_code} (1-{len(unique_locations) + 1}): ").strip()
        
        try:
            choice_num = int(choice)
            if 1 <= choice_num <= len(unique_locations):
                selected_location = unique_locations[choice_num - 1]
                existing_locations[set_upper] = selected_location
                return selected_location
            elif choice_num == len(unique_locations) + 1:
                new_location = input("Enter new location name: ").strip()
                if new_location:
                    existing_locations[set_upper] = new_location
                    return new_location
                else:
                    print("Location name cannot be empty.")
            else:
                print(f"Please enter a number between 1 and {len(unique_locations) + 1}.")
        except ValueError:
            print("Please enter a valid number.")


def format_high_value_reminder(items: List[Dict[str, Any]]) -> str:
    """Generate a reminder message for high-value cards that may be in binders."""
    high_value_items = [item for item in items if item.get('price', 0) >= 10.0]
    
    if not high_value_items:
        return ""
    
    reminder = f"\n💰 HIGH VALUE ALERT: {len(high_value_items)} cards worth $10+ may be in binders:\n"
    for item in sorted(high_value_items, key=lambda x: x.get('price', 0), reverse=True):
        price = item.get('price', 0)
        name = item.get('name', 'N/A')
        set_code = item.get('set', 'N/A')
        order_id = item.get('order_id', 'N/A')
        reminder += f"  • ${price:.2f} - {name} [{set_code}] (Order: {order_id})\n"
    
    return reminder


def get_scryfall_image_uri(card_name: str, set_code: str,
                           collector_number: str) -> str:
    """Return a normal-sized image URL from Scryfall (or 'N/A' if not found)."""
    # Scryfall API requests a 50-100ms delay between requests
    time.sleep(0.1)

    # First try exact set / collector lookup
    try:
        url = f"{SCRYFALL_API_BASE_URL}/cards/{set_code.lower()}/{collector_number}"
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return r.json().get("image_uris", {}).get("normal", "N/A")
    except requests.exceptions.RequestException:
        pass  # Fallback to fuzzy search

    # Fallback: fuzzy name search
    try:
        url = f"{SCRYFALL_API_BASE_URL}/cards/named"
        r = requests.get(url, params={"fuzzy": card_name}, timeout=10)
        r.raise_for_status()
        return r.json().get("image_uris", {}).get("normal", "N/A")
    except requests.exceptions.RequestException:
        logging.warning(
            f"Could not find image for: {card_name} [{set_code}-{collector_number}]"
        )
        return "N/A"


def get_order_details(order_id: str,
                      headers: Dict[str, str]) -> Optional[Dict[str, Any]]:
    """Fetch the full order details for *order_id*."""
    try:
        url = f"{MANAPOOL_API_BASE_URL}/seller/orders/{order_id}"
        r = requests.get(url, headers=headers, timeout=15)
        r.raise_for_status()
        return r.json().get("order")
    except (requests.exceptions.RequestException, ValueError, KeyError) as exc:
        logging.error(f"Could not fetch details for order {order_id}: {exc}")
        return None


def get_user_html_report_choice(settings: Dict[str, Any]) -> bool:
    """Prompt the user to enable/disable HTML report generation."""
    # Check saved settings - use them directly if they exist
    if settings.get("generate_html_report") is not None:
        logging.info(f"Using saved HTML report preference: {'Enabled' if settings['generate_html_report'] else 'Disabled'}")
        return settings["generate_html_report"]
    
    # If not set in settings, prompt user for initial setup
    while True:
        print("\n🎨 HTML Report Options:")
        print("  Generate visual HTML report? This will:")
        print("  • Create an interactive HTML file with card images")
        print("  • Show cards organized by location")
        print("  • Include progress tracking checkboxes")
        print("  • Open automatically in your browser")
        print("  • Allow quick cleanup when finished")
        print("  (This choice will be saved for future runs)")
        choice = input("Generate HTML report? (Y/n): ").strip().lower()
        
        if choice in ["", "y", "yes"]:
            return True
        elif choice in ["n", "no"]:
            return False
        else:
            print("Please enter 'y' for yes or 'n' for no.")


def get_user_filter_choice(settings: Dict[str, Any]) -> str:
    """Prompt the user to select an order filter and return their choice."""
    print("\n" + "=" * 40)
    print("  ManaPool Order Retrieval")
    print("=" * 40)
    
    while True:
        print("\nPlease select which orders to retrieve:")
        print("  1: Not Shipped (Default)")
        print("  2: Shipped Only")
        print("  3: All Orders")
        if settings.get("download_images") is not None or settings.get("generate_html_report") is not None:
            print("  (Type 'reset' to change saved preferences)")
        choice = input("Enter your choice (1-3): ").strip() or "1"
        if choice.lower() == "reset":
            settings["download_images"] = None
            settings["generate_html_report"] = None
            print("Preferences reset. You'll be prompted to set them again.")
            continue
        if choice in ["1", "2", "3"]:
            return choice
        print("\nInvalid choice. Please enter 1, 2, or 3.")


def get_user_image_download_choice(settings: Dict[str, Any]) -> bool:
    """Prompt the user to enable/disable image downloading."""
    # Check environment variable first
    download_images_env = os.getenv("DOWNLOAD_IMAGES", "").lower()
    if download_images_env in ["true", "1", "yes"]:
        return True
    elif download_images_env in ["false", "0", "no"]:
        return False
    
    # Check saved settings - use them directly if they exist
    if settings.get("download_images") is not None:
        logging.info(f"Using saved image download preference: {'Enabled' if settings['download_images'] else 'Disabled'}")
        return settings["download_images"]
    
    # If not set in environment or settings, prompt user for initial setup
    while True:
        print("\n📸 Image Download Options:")
        print("  Download card images locally? This will:")
        print("  • Create an 'images' folder with card images")
        print("  • CSV will always contain Scryfall URLs")
        print("  • HTML report will use local images for offline viewing")
        print("  • Skip downloading images that already exist")
        print("  • Take additional time for initial downloads")
        print("  (This choice will be saved for future runs)")
        choice = input("Download images? (y/N): ").strip().lower()
        
        if choice in ["", "n", "no"]:
            return False
        elif choice in ["y", "yes"]:
            return True
        else:
            print("Please enter 'y' for yes or 'n' for no.")


def generate_html_report(items: List[Dict[str, Any]], output_filename: str, filter_description: str, timestamp: str, download_images: bool) -> str:
    """Generate an HTML report with embedded images and interactive features."""
    html_filename = output_filename.replace('.csv', '.html')
    
    # Group items by location for better organization
    items_by_location = {}
    for item in items:
        location = item.get('location', 'Unknown')
        if location not in items_by_location:
            items_by_location[location] = []
        items_by_location[location].append(item)
    
    # Count statistics
    total_items = len(items)
    total_value = sum(item.get('price', 0) for item in items)
    unique_orders = len(set(item.get('order_id', '') for item in items))
    high_value_count = len([item for item in items if item.get('price', 0) >= 10.0])
    
    html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ManaPool Fulfillment Report - {filter_description}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
            color: #333;
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            border-radius: 10px;
            margin-bottom: 30px;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }}
        .header h1 {{
            margin: 0 0 10px 0;
            font-size: 2.5em;
        }}
        .header p {{
            margin: 5px 0;
            font-size: 1.1em;
            opacity: 0.9;
        }}
        .stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        .stat-card {{
            background: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            border-left: 4px solid #667eea;
        }}
        .stat-number {{
            font-size: 2em;
            font-weight: bold;
            color: #667eea;
            margin-bottom: 5px;
        }}
        .stat-label {{
            color: #666;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        .location-section {{
            background: white;
            margin-bottom: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            overflow: hidden;
        }}
        .location-header {{
            background: #667eea;
            color: white;
            padding: 20px;
            font-size: 1.3em;
            font-weight: bold;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .location-checkbox {{
            margin-right: 10px;
            transform: scale(1.2);
        }}
        .cards-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
            gap: 20px;
            padding: 20px;
        }}
        .card-item {{
            border: 1px solid #ddd;
            border-radius: 8px;
            overflow: hidden;
            background: #fafafa;
            transition: transform 0.2s, box-shadow 0.2s;
        }}
        .card-item:hover {{
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        }}
        .card-image {{
            width: 100%;
            height: 300px;
            background: #f0f0f0;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #888;
            font-style: italic;
            overflow: hidden;
            border-radius: 8px 8px 0 0;
        }}
        .card-image img {{
            max-width: 100%;
            max-height: 100%;
            object-fit: contain;
            border-radius: 8px;
        }}
        .card-details {{
            padding: 15px;
        }}
        .card-header {{
            display: flex;
            align-items: center;
            margin-bottom: 8px;
        }}
        .card-checkbox {{
            margin-right: 8px;
            transform: scale(1.1);
        }}
        .card-name {{
            font-weight: bold;
            font-size: 1.1em;
            color: #333;
            cursor: pointer;
        }}
        .card-meta {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
            font-size: 0.9em;
            color: #666;
            margin-bottom: 10px;
        }}
        .card-price {{
            font-size: 1.2em;
            font-weight: bold;
            color: #2c5aa0;
            text-align: center;
            padding: 8px;
            background: #e8f2ff;
            border-radius: 4px;
        }}
        .high-value {{
            border-left: 4px solid #ff6b6b;
        }}
        .high-value .card-price {{
            background: #ffe8e8;
            color: #c92a2a;
        }}
        .order-info {{
            background: #f8f9fa;
            padding: 8px;
            border-radius: 4px;
            margin-top: 8px;
            font-size: 0.85em;
            color: #555;
        }}
        .no-image {{
            background: #f0f0f0;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #888;
            font-style: italic;
            height: 200px;
        }}
        @media (max-width: 768px) {{
            .cards-grid {{
                grid-template-columns: 1fr;
            }}
            .stats {{
                grid-template-columns: repeat(2, 1fr);
            }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📦 ManaPool Fulfillment Report</h1>
        <p><strong>{filter_description}</strong> Orders</p>
        <p>Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
    </div>
    
    <div class="stats">
        <div class="stat-card">
            <div class="stat-number">{total_items}</div>
            <div class="stat-label">Total Items</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">{unique_orders}</div>
            <div class="stat-label">Unique Orders</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">${total_value:.2f}</div>
            <div class="stat-label">Total Value</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">{high_value_count}</div>
            <div class="stat-label">High Value Cards ($10+)</div>
        </div>
    </div>
"""
    
    # Generate location sections
    for location, location_items in items_by_location.items():
        location_total = sum(item.get('price', 0) for item in location_items)
        location_checkbox_id = f"location_{location.replace(' ', '_').replace('/', '_')}"
        safe_location = html.escape(location)
        
        html_content += f"""
    <div class="location-section" id="{html.escape(location_checkbox_id)}_section">
        <div class="location-header">
            <div>
                <input type="checkbox" id="{html.escape(location_checkbox_id)}" class="location-checkbox" onchange="toggleLocationCards('{html.escape(location_checkbox_id)}')">
                <label for="{html.escape(location_checkbox_id)}">📍 {safe_location} ({len(location_items)} items - ${location_total:.2f})</label>
            </div>
        </div>
        <div class="cards-grid">
"""
        
        for item in location_items:
            card_name = item.get('name', 'Unknown Card')
            set_code = item.get('set', 'N/A')
            collector_number = item.get('number', 'N/A')
            condition = item.get('condition', 'N/A')
            finish = item.get('finish', 'N/A')
            price = item.get('price', 0)
            quantity = item.get('quantity', 1)
            order_id = item.get('order_id', 'N/A')
            order_label = item.get('order_label', 'N/A')
            scryfall_uri = item.get('scryfall_image_uri', 'N/A')
            local_image_path = item.get('local_image_path', None)
            
            # Check if it's a high-value card
            high_value_class = 'high-value' if price >= 10.0 else ''
            
            # Handle image display - prefer local images if available, fallback to Scryfall
            if download_images and local_image_path and local_image_path != 'N/A' and os.path.exists(local_image_path):
                # Use local image
                relative_path = os.path.relpath(local_image_path).replace('\\', '/')
                safe_card_name = html.escape(card_name)
                image_html = f'<img src="{html.escape(relative_path)}" alt="{safe_card_name}" onerror="this.src=\'{html.escape(scryfall_uri)}\'; this.onerror=null;">'
            elif scryfall_uri != 'N/A':
                # Use Scryfall URL
                safe_card_name = html.escape(card_name)
                image_html = f'<img src="{html.escape(scryfall_uri)}" alt="{safe_card_name}" onerror="this.parentElement.innerHTML=\'No Image Available\';">'
            else:
                image_html = 'No Image Available'
            
            # Generate unique checkbox ID
            checkbox_id = f"card_{order_id}_{collector_number}_{card_name.replace(' ', '_').replace('/', '_')}"
            
            # Escape all text content for security
            safe_card_name = html.escape(card_name)
            safe_set_code = html.escape(set_code)
            safe_collector_number = html.escape(collector_number)
            safe_condition = html.escape(condition)
            safe_finish = html.escape(finish)
            safe_order_label = html.escape(order_label)
            safe_order_id = html.escape(order_id)
            
            html_content += f"""
            <div class="card-item {high_value_class}" id="{html.escape(checkbox_id)}_container">
                <div class="card-image">
                    {image_html}
                </div>
                <div class="card-details">
                    <div class="card-header">
                        <input type="checkbox" id="{html.escape(checkbox_id)}" class="card-checkbox" onchange="updateProgress()">
                        <label for="{html.escape(checkbox_id)}" class="card-name">{safe_card_name}</label>
                    </div>
                    <div class="card-meta">
                        <span><strong>Set:</strong> {safe_set_code}</span>
                        <span><strong>Number:</strong> {safe_collector_number}</span>
                        <span><strong>Condition:</strong> {safe_condition}</span>
                        <span><strong>Finish:</strong> {safe_finish}</span>
                    </div>
                    <div class="card-price">${price:.2f} × {quantity}</div>
                    <div class="order-info">
                        <strong>Order:</strong> {safe_order_label} ({safe_order_id})<br>
                        <a href="https://manapool.com/seller/orders/{safe_order_id}" target="_blank" style="color: #667eea; text-decoration: none; font-weight: bold;">📋 View Order Page →</a>
                    </div>
                </div>
            </div>
"""
        
        html_content += """
        </div>
    </div>
"""
    
    html_content += f"""
    
    <div style="position: fixed; bottom: 20px; right: 20px; background: white; padding: 15px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.2); z-index: 1000;">
        <div style="margin-bottom: 10px; font-weight: bold;">Progress: <span id="progress-text">0/0</span></div>
        <button onclick="showOrderView()" style="background: #007bff; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 14px; margin-bottom: 10px; width: 100%;">📦 View by Orders</button><br>
        <button onclick="finishAndCleanup()" style="background: #28a745; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 16px; width: 100%;">✅ Finished - Close & Delete</button>
    </div>
    
    <script>
        function updateProgress() {{
            const allCards = document.querySelectorAll('.card-checkbox');
            const checkedCards = document.querySelectorAll('.card-checkbox:checked');
            const progressText = document.getElementById('progress-text');
            progressText.textContent = `${{checkedCards.length}}/${{allCards.length}}`;
        }}
        
        function toggleLocationCards(locationId) {{
            const locationCheckbox = document.getElementById(locationId);
            const locationSection = document.getElementById(locationId + '_section');
            const cardCheckboxes = locationSection.querySelectorAll('.card-checkbox');
            
            cardCheckboxes.forEach(checkbox => {{
                checkbox.checked = locationCheckbox.checked;
            }});
            
            updateProgress();
        }}
        
        let originalBodyContent = '';
        let isOrderView = false;
        
        function showOrderView() {{
            if (!isOrderView) {{
                // Store original content
                originalBodyContent = document.body.innerHTML;
            }}
            
            // Create order-grouped data structure
            const orderGroups = {{}};
            const allCards = document.querySelectorAll('.card-item');
            
            allCards.forEach(card => {{
                const orderInfo = card.querySelector('.order-info');
                if (!orderInfo) return;
                
                const orderText = orderInfo.textContent;
                const orderIdMatch = orderText.match(/\\(([^)]+)\\)/);
                const orderLabelMatch = orderText.match(/Order:\\s*([^(]+)/);
                
                if (orderIdMatch && orderLabelMatch) {{
                    const orderId = orderIdMatch[1].trim();
                    const orderLabel = orderLabelMatch[1].trim();
                    
                    if (!orderGroups[orderId]) {{
                        orderGroups[orderId] = {{
                            label: orderLabel,
                            cards: [],
                            totalValue: 0,
                            totalQuantity: 0
                        }};
                    }}
                    
                    const cardName = card.querySelector('.card-name').textContent;
                    const cardMeta = card.querySelector('.card-meta');
                    const priceElement = card.querySelector('.card-price');
                    const price = priceElement.textContent;
                    
                    // Find location
                    let location = 'Unknown';
                    const locationSection = card.closest('.location-section');
                    if (locationSection) {{
                        const locationLabel = locationSection.querySelector('.location-header label');
                        if (locationLabel) {{
                            const locationMatch = locationLabel.textContent.match(/📍\\s*([^(]+)/);
                            if (locationMatch) {{
                                location = locationMatch[1].trim();
                            }}
                        }}
                    }}
                    
                    // Extract price and quantity
                    const priceMatch = price.match(/\\$([0-9.]+)(?:\\s*×\\s*(\\d+))?/);
                    let itemPrice = 0;
                    let quantity = 1;
                    if (priceMatch) {{
                        itemPrice = parseFloat(priceMatch[1]);
                        quantity = priceMatch[2] ? parseInt(priceMatch[2]) : 1;
                    }}
                    
                    orderGroups[orderId].cards.push({{
                        name: cardName,
                        meta: cardMeta ? cardMeta.innerHTML : '',
                        price: price,
                        location: location,
                        element: card.cloneNode(true),
                        unitPrice: itemPrice,
                        quantity: quantity
                    }});
                    
                    orderGroups[orderId].totalValue += itemPrice * quantity;
                    orderGroups[orderId].totalQuantity += quantity;
                }}
            }});
            
            // Create order view HTML
            let orderViewHTML = `
                <div class="header">
                    <h1>📦 Orders Overview</h1>
                    <p>Cards grouped by individual orders</p>
                    <button onclick="showLocationView()" style="background: #6c757d; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; margin-top: 10px;">← Back to Location View</button>
                </div>
            `;
            
            // Sort orders by total value (highest first)
            const sortedOrders = Object.entries(orderGroups).sort((a, b) => b[1].totalValue - a[1].totalValue);
            
            sortedOrders.forEach(([orderId, orderData]) => {{
                const safeOrderId = orderId.replace(/[^a-zA-Z0-9]/g, '_');
                orderViewHTML += `
                    <div class="location-section">
                        <div class="location-header">
                            <div>
                                <input type="checkbox" id="order_${{safeOrderId}}_checkbox" class="location-checkbox" onchange="toggleOrderCards('${{safeOrderId}}')">
                                <label for="order_${{safeOrderId}}_checkbox">📋 Order ${{orderData.label}} (${{orderData.totalQuantity}} items - $$${{orderData.totalValue.toFixed(2)}})</label>
                            </div>
                            <div>
                                <a href="https://manapool.com/seller/orders/${{orderId}}" target="_blank" style="color: white; text-decoration: none; font-weight: bold;">📋 View Order Page →</a>
                            </div>
                        </div>
                        <div class="cards-grid" id="order_${{safeOrderId}}_cards">
                `;
                
                orderData.cards.forEach((card, cardIndex) => {{
                    const cardCheckboxId = `order_${{safeOrderId}}_card_${{cardIndex}}`;
                    orderViewHTML += `
                        <div class="card-item">
                            <div class="card-image">
                                ${{card.element.querySelector('.card-image').innerHTML}}
                            </div>
                            <div class="card-details">
                                <div class="card-header">
                                    <input type="checkbox" id="${{cardCheckboxId}}" class="card-checkbox" onchange="updateProgress()">
                                    <label for="${{cardCheckboxId}}" class="card-name">${{card.name}}</label>
                                </div>
                                <div class="card-meta">
                                    ${{card.meta}}
                                </div>
                                <div class="card-price">${{card.price}}</div>
                                <div style="background: #e8f2ff; padding: 8px; margin-top: 8px; border-radius: 4px; font-size: 0.85em;">
                                    <strong>📍 Location:</strong> ${{card.location}}
                                </div>
                                <div class="order-info">
                                    <strong>Order:</strong> ${{orderData.label}} (${{orderId}})<br>
                                    <a href="https://manapool.com/seller/orders/${{orderId}}" target="_blank" style="color: #667eea; text-decoration: none; font-weight: bold;">📋 View Order Page →</a>
                                </div>
                            </div>
                        </div>
                    `;
                }});
                
                orderViewHTML += `
                        </div>
                    </div>
                `;
            }});
            
            // Replace current content
            document.body.innerHTML = orderViewHTML + `
                <div style="position: fixed; bottom: 20px; right: 20px; background: white; padding: 15px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.2); z-index: 1000;">
                    <div style="margin-bottom: 10px; font-weight: bold;">Progress: <span id="progress-text">0/0</span></div>
                    <button onclick="showLocationView()" style="background: #6c757d; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 14px; margin-bottom: 10px; width: 100%;">← Back to Location View</button><br>
                    <button onclick="finishAndCleanup()" style="background: #28a745; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; font-size: 16px; width: 100%;">✅ Finished - Close & Delete</button>
                </div>
            `;
            
            isOrderView = true;
            updateProgress();
        }}
        
        function showLocationView() {{
            if (isOrderView && originalBodyContent) {{
                document.body.innerHTML = originalBodyContent;
                isOrderView = false;
                // Reinitialize progress counter
                updateProgress();
            }} else {{
                location.reload();
            }}
        }}
        
        function toggleOrderCards(safeOrderId) {{
            const orderCheckbox = document.getElementById('order_' + safeOrderId + '_checkbox');
            const orderSection = document.getElementById('order_' + safeOrderId + '_cards');
            if (orderSection) {{
                const cardCheckboxes = orderSection.querySelectorAll('.card-checkbox');
                cardCheckboxes.forEach(checkbox => {{
                    checkbox.checked = orderCheckbox.checked;
                }});
                updateProgress();
            }}
        }}

        function finishAndCleanup() {{
            if (confirm('Are you sure you want to close this report and delete the HTML file?')) {{
                // Send request to delete the file
                fetch(window.location.href, {{
                    method: 'DELETE'
                }}).catch(() => {{
                    // Ignore errors - file might not be deletable
                }});
                
                // Close the window/tab
                window.close();
                
                // If window.close() doesn't work (some browsers), show message
                setTimeout(() => {{
                    alert('Please close this tab manually. The report file has been marked for deletion.');
                }}, 100);
            }}
        }}
        
        // Initialize progress counter
        window.addEventListener('load', updateProgress);
    </script>
</body>
</html>
"""
    
    # Write the HTML file
    try:
        with open(html_filename, 'w', encoding='utf-8') as f:
            f.write(html_content)
        logging.info(f"Successfully generated HTML report: {html_filename}")
        return html_filename
    except Exception as exc:
        logging.error(f"Failed to generate HTML report: {exc}")
        return ""


def main() -> None:
    # --- Logging Setup ---
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s - %(levelname)s - %(message)s",
                        handlers=[
                            logging.FileHandler(LOG_FILE, mode='w'),
                            logging.StreamHandler(sys.stdout)
                        ])

    # --- Load User Settings and Get Choices ---
    user_settings = load_user_settings(SETTINGS_FILE)
    
    filter_choice = get_user_filter_choice(user_settings)
    download_images = get_user_image_download_choice(user_settings)
    generate_html = get_user_html_report_choice(user_settings)
    
    # Save updated settings
    user_settings["download_images"] = download_images
    user_settings["generate_html_report"] = generate_html
    save_user_settings(SETTINGS_FILE, user_settings)
    
    set_locations = load_set_locations(LOCATIONS_FILE)
    if set_locations:
        logging.info(
            f"Successfully loaded {len(set_locations)} set locations.")
    
    # Load image cache if downloading images
    image_cache = {}
    if download_images:
        image_cache = load_image_cache(IMAGE_CACHE_FILE)
        logging.info(f"Image download enabled. Loaded {len(image_cache)} cached images.")

    # --- Credentials ---
    # Read credentials securely from environment variables
    email = os.getenv("MANAPOOL_EMAIL")
    token = os.getenv("MANAPOOL_API_KEY")

    if not email or not token:
        logging.error(
            "MANAPOOL_EMAIL or MANAPOOL_API_KEY not found in .env file.")
        logging.error("Please create a .env file and add your credentials.")
        return

    headers = {
        "X-ManaPool-Email": email,
        "X-ManaPool-Access-Token": token,
    }

    # --- Fetch All Orders ---
    logging.info("Fetching all orders from ManaPool…")
    try:
        resp = requests.get(f"{MANAPOOL_API_BASE_URL}/seller/orders",
                            headers=headers,
                            timeout=20)
        resp.raise_for_status()
        data = resp.json()
        orders_list: List[Dict[
            str, Any]] = data.get("data") or data.get("orders") or data
    except requests.exceptions.RequestException as exc:
        logging.error(f"API communication error: {exc}")
        if exc.response is not None:
            logging.error(f"Response content: {exc.response.text}")
        return
    except (ValueError, KeyError) as exc:
        logging.error(f"Error parsing API response: {exc}")
        return

    logging.info(f"Total orders returned from API: {len(orders_list)}")
    
    timestamp = datetime.now().strftime("%Y%m%d%H%M")

    # --- Filter Orders Based on User Choice ---
    if filter_choice == "1":
        filtered_orders = [
            o for o in orders_list if o.get(FULFILLMENT_FIELD) != SHIPPED_VALUE
        ]
        filter_description = "Not Shipped"
        output_filename = f"{BASE_OUTPUT_CSV_NAME}_not_shipped_{timestamp}.csv"
    elif filter_choice == "2":
        filtered_orders = [
            o for o in orders_list if o.get(FULFILLMENT_FIELD) == SHIPPED_VALUE
        ]
        filter_description = "Shipped"
        output_filename = f"{BASE_OUTPUT_CSV_NAME}_shipped_{timestamp}.csv"
    else:  # choice == "3"
        filtered_orders = orders_list
        filter_description = "All"
        output_filename = f"{BASE_OUTPUT_CSV_NAME}_all_{timestamp}.csv"

    logging.info(
        f"Filtering for '{filter_description}' orders. Found {len(filtered_orders)} matching orders."
    )

    if not filtered_orders:
        logging.info("No orders match the selected filter — exiting.")
        return

    # --- Process Orders and Line Items ---
    items: List[Dict[str, Any]] = []
    total_to_process = len(filtered_orders)
    new_sets_found: Set[str] = set()
    logging.info(f"Fetching details for {total_to_process} orders...")

    for i, summary in enumerate(filtered_orders):
        order_id = summary.get("id")
        if not order_id:
            continue

        logging.info(
            f"Processing order {i+1}/{total_to_process} (ID: {order_id})")
        details = get_order_details(order_id, headers)
        if not details:
            continue

        for item in details.get("items", []):
            product = item.get("product", {})
            single = product.get("single", {})

            card_name = single.get("name", "N/A")
            set_code = single.get("set", "N/A")
            collector_number = single.get("number", "N/A")

            # Get location with interactive assignment for new sets
            location = get_location_for_set(set_code, set_locations, new_sets_found)

            image_uri = get_scryfall_image_uri(card_name, set_code,
                                               collector_number)

            # Handle image downloading if enabled (for HTML use)
            local_image_path = None
            if download_images and image_uri != "N/A":
                image_key = generate_image_key(card_name, set_code, collector_number)
                
                # Check cache first
                if image_key in image_cache:
                    local_image_path = image_cache[image_key]
                    # Verify file still exists
                    if not os.path.exists(local_image_path):
                        # File was deleted, re-download
                        local_image_path = download_card_image(image_uri, image_key, IMAGES_DIR)
                        image_cache[image_key] = local_image_path
                else:
                    # Download new image
                    local_image_path = download_card_image(image_uri, image_key, IMAGES_DIR)
                    image_cache[image_key] = local_image_path

            items.append({
                "order_id": order_id,
                "order_label": details.get("label", "N/A"),
                "location": location,  # Added location column
                "quantity": item.get("quantity", 0),
                "name": card_name,
                "set": set_code,
                "number": collector_number,
                "condition": single.get("condition_id", "N/A"),
                "finish": single.get("finish_id", "N/A"),
                "price": item.get("price_cents", 0) / 100.0,
                "tcgplayer_sku": product.get("tcgplayer_sku", "N/A"),
                "scryfall_image_uri": image_uri,  # Always store original Scryfall URL for CSV
                "local_image_path": local_image_path,  # Store local path separately for HTML
            })

    if not items:
        logging.info(
            "No line items found in the processed orders — nothing exported.")
        return

    # --- Save Updated Locations ---
    if new_sets_found:
        save_set_locations(LOCATIONS_FILE, {k.upper(): v for k, v in set_locations.items()})
    
    # --- Save Updated Image Cache ---
    if download_images and image_cache:
        save_image_cache(IMAGE_CACHE_FILE, image_cache)
        logging.info(f"Saved {len(image_cache)} images to cache.")

    # --- High Value Card Reminder ---
    high_value_reminder = format_high_value_reminder(items)
    if high_value_reminder:
        print(high_value_reminder)
        logging.info(f"Found {len([i for i in items if i.get('price', 0) >= 10.0])} high-value cards (>=$10)")

    # --- Export to CSV ---
    df = pd.DataFrame(items)
    # Reorder columns to have location near the front
    cols = [
        'order_id', 'order_label', 'location', 'quantity', 'name', 'set',
        'number', 'condition', 'finish', 'price', 'tcgplayer_sku',
        'scryfall_image_uri'
    ]
    df = df[cols]
    
    logging.info('Sorting items for output...')
    
    df['location_total_qty'] = df.groupby('location')['quantity'].transform('sum')
    df['set_total_qty'] = df.groupby(['location', 'set'])['quantity'].transform('sum')
    
    # Sort by:
    # 1. Location total quantity (desc) - most overall quantity in a location first
    # 2. Location name (asc) - for consistent tie-breaking
    # 3. Set total quantity (desc) - most quantity per set within location first  
    # 4. Set name (asc) - for consistent tie-breaking
    # 5. Card name (asc) - alphabetical card names
    df.sort_values(by=['location_total_qty', 'location', 'set_total_qty', 'set', 'name'],
                   ascending=[False, True, False, True, True], inplace=True)
    
    
    
    df.drop(columns=['location_total_qty', 'set_total_qty'], inplace=True)
    



    try:
        df.to_csv(output_filename, index=False)
        logging.info(
            f"Successfully exported {len(items)} items to {output_filename}")
    except Exception as exc:
        logging.error(f"Failed to write CSV file: {exc}")

    # --- Generate HTML Report (if enabled) ---
    html_filename = ""
    if generate_html:
        logging.info("Generating HTML report...")
        html_filename = generate_html_report(items, output_filename, filter_description, timestamp, download_images)
        
        if html_filename:
            try:
                # Convert to absolute path for browser
                abs_html_path = os.path.abspath(html_filename)
                webbrowser.open(f"file://{abs_html_path}")
                logging.info(f"Opened HTML report in default browser: {html_filename}")
                print(f"\n🌐 HTML report opened in your default browser!")
                print(f"📄 CSV file: {output_filename}")
                print(f"🎨 HTML report: {html_filename}")
            except Exception as exc:
                logging.warning(f"Could not open browser automatically: {exc}")
                print(f"\n📄 Files generated:")
                print(f"  • CSV: {output_filename}")
                print(f"  • HTML: {html_filename}")
                print(f"💡 Open {html_filename} in your browser to view the visual report!")
    else:
        print(f"\n📄 CSV file generated: {output_filename}")
        print(f"💡 HTML report generation was disabled. Enable it in settings if you want visual reports.")

    logging.info("Order retrieval process complete.")


if __name__ == "__main__":
    main()
