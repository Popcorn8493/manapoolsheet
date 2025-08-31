#!/usr/bin/env python3

import html
import json
import logging
import os
import re
import sys
import time
import webbrowser
from datetime import datetime
from typing import Any, Dict, List, Optional, Set

import pandas as pd
import requests
from dotenv import load_dotenv

# --- Configuration ---

# Load environment variables from a .env file
load_dotenv()

LOG_FILE = "order_processing_log.txt"
BASE_OUTPUT_CSV_NAME = "fulfillment_list"
LOCATIONS_FILE = "locations.json"  # File to map MTG sets to physical storage locations
IMAGES_DIR = "images"  # Directory to store downloaded card images
IMAGE_CACHE_FILE = "image_cache.json"  # Cache file to track downloaded images
SETTINGS_FILE = "settings.json"  # User preferences file

# Data directory structure
DATA_DIR = "data"
CSV_DIR = os.path.join(DATA_DIR, "csv")
HTML_DIR = os.path.join(DATA_DIR, "html")

# API Configuration
ORDERS_API_BASE_URL = "https://api.cardmarket.com/v1"  # Replace with your marketplace API
IMAGES_API_BASE_URL = "https://api.scryfall.com"  # Scryfall API for card images
FULFILLMENT_FIELD = "fulfillment_status"
SHIPPED_VALUE = "shipped"


# --- Helper Functions ---

def load_set_locations(filename: str) -> Dict[str, str]:
	"""Loads the MTG set-to-location mapping from a JSON file."""
	try:
		with open(filename, "r", encoding="utf-8") as f:
			locations = json.load(f)
			# Ensure all keys are uppercase for consistent matching
			return {k.upper(): v for k, v in locations.items()}
	except FileNotFoundError:
		logging.warning(f"Location file not found: {filename}. Location data will be missing.")
		return {}
	except json.JSONDecodeError:
		logging.error(f"Could not parse {filename}. Please ensure it is valid JSON.")
		return {}


def save_set_locations(filename: str, locations: Dict[str, str]) -> None:
	"""Saves the MTG set-to-location mapping to a JSON file."""
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
				"download_images":      None,
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
	"""Generate a unique key for caching downloaded card images."""
	# Create a consistent key that uniquely identifies a card
	safe_name = card_name.replace(' ', '_').replace('/', '_')
	return f"{set_code.upper()}_{collector_number}_{safe_name}"


def download_card_image(image_url: str, image_key: str, images_dir: str) -> str:
	"""Download a card image and return the local file path."""
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
	"""Get location for an MTG set, prompting user for new sets."""
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
	
	print(f"\n📦 New MTG set found: {set_code}")
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


def get_scryfall_image_uri(card_name: str, set_code: str, collector_number: str) -> str:
	"""Return a card image URL from Scryfall API."""
	# Scryfall API requests a 50-100ms delay between requests
	time.sleep(0.1)
	
	# First try exact set / collector lookup
	try:
		url = f"{IMAGES_API_BASE_URL}/cards/{set_code.lower()}/{collector_number}"
		r = requests.get(url, timeout=10)
		r.raise_for_status()
		return r.json().get("image_uris", {}).get("normal", "N/A")
	except requests.exceptions.RequestException:
		pass  # Fallback to fuzzy search
	
	# Fallback: fuzzy name search
	try:
		url = f"{IMAGES_API_BASE_URL}/cards/named"
		r = requests.get(url, params={"fuzzy": card_name}, timeout=10)
		r.raise_for_status()
		return r.json().get("image_uris", {}).get("normal", "N/A")
	except requests.exceptions.RequestException:
		logging.warning(f"Could not find image for: {card_name} [{set_code}-{collector_number}]")
		return "N/A"


def get_order_details(order_id: str, headers: Dict[str, str]) -> Optional[Dict[str, Any]]:
	"""Fetch the full order details for order_id."""
	try:
		url = f"{ORDERS_API_BASE_URL}/seller/orders/{order_id}"
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
		logging.info(
				f"Using saved HTML report preference: {'Enabled' if settings['generate_html_report'] else 'Disabled'}")
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
	print("  MTG Order Fulfillment Tool")
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
		logging.info(
				f"Using saved image download preference: {'Enabled' if settings['download_images'] else 'Disabled'}")
		return settings["download_images"]
	
	# If not set in environment or settings, prompt user for initial setup
	while True:
		print("\n📸 Image Download Options:")
		print("  Download card images locally? This will:")
		print("  • Create an 'images' folder with card images")
		print("  • CSV will always contain image URLs")
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


def generate_html_report(items: List[Dict[str, Any]], output_filename: str,
                         filter_description: str, timestamp: str, download_images: bool) -> str:
	"""Generate an HTML report with embedded images and interactive features."""
	# Extract filename from path and create HTML filename in HTML directory
	csv_filename = os.path.basename(output_filename)
	html_filename = os.path.join(HTML_DIR, csv_filename.replace('.csv', '.html'))
	
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
	
	# HTML template with interactive features
	html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Order Management Report - {filter_description}</title>
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
        }}
        .location-section {{
            background: white;
            margin-bottom: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }}
        .location-header {{
            background: #667eea;
            color: white;
            padding: 20px;
            font-size: 1.3em;
            font-weight: bold;
        }}
        .items-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
            gap: 20px;
            padding: 20px;
        }}
        .item-card {{
            border: 1px solid #ddd;
            border-radius: 8px;
            overflow: hidden;
            background: #fafafa;
        }}
        .item-details {{
            padding: 15px;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📦 MTG Order Fulfillment Report</h1>
        <p><strong>{filter_description}</strong> Orders</p>
        <p>Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
    </div>
    
    <div class="stats">
        <div class="stat-card">
            <div style="font-size: 2em; font-weight: bold; color: #667eea;">{total_items}</div>
            <div>Total Items</div>
        </div>
        <div class="stat-card">
            <div style="font-size: 2em; font-weight: bold; color: #667eea;">{unique_orders}</div>
            <div>Unique Orders</div>
        </div>
        <div class="stat-card">
            <div style="font-size: 2em; font-weight: bold; color: #667eea;">${total_value:.2f}</div>
            <div>Total Value</div>
        </div>
        <div class="stat-card">
            <div style="font-size: 2em; font-weight: bold; color: #667eea;">{high_value_count}</div>
            <div>High Value Cards ($10+)</div>
        </div>
    </div>
"""
	
	# Add location sections (simplified)
	for location, location_items in items_by_location.items():
		location_total = sum(item.get('price', 0) for item in location_items)
		safe_location = html.escape(location)
		
		html_content += f"""
    <div class="location-section">
        <div class="location-header">
            📍 {safe_location} ({len(location_items)} items - ${location_total:.2f})
        </div>
        <div class="items-grid">
"""
		
		for item in location_items:
			card_name = html.escape(item.get('name', 'Unknown Card'))
			set_code = html.escape(item.get('set', 'N/A'))
			price = item.get('price', 0)
			quantity = item.get('quantity', 1)
			order_id = html.escape(item.get('order_id', 'N/A'))
			
			html_content += f"""
            <div class="item-card">
                <div class="item-details">
                    <div style="font-weight: bold; font-size: 1.1em;">{card_name}</div>
                    <div style="margin: 8px 0;">
                        <strong>Set:</strong> {set_code}<br>
                        <strong>Price:</strong> ${price:.2f} × {quantity}
                    </div>
                    <div style="background: #f8f9fa; padding: 8px; border-radius: 4px; font-size: 0.85em;">
                        <strong>Order:</strong> {order_id}
                    </div>
                </div>
            </div>
"""
		
		html_content += """
        </div>
    </div>
"""
	
	html_content += """
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
	"""Main function for MTG order processing and fulfillment list generation."""
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
		logging.info(f"Successfully loaded {len(set_locations)} set locations.")
	
	# Load image cache if downloading images
	image_cache = {}
	if download_images:
		image_cache = load_image_cache(IMAGE_CACHE_FILE)
		logging.info(f"Image download enabled. Loaded {len(image_cache)} cached images.")
	
	# --- Credentials ---
	# Read credentials securely from environment variables
	email = os.getenv("MARKETPLACE_EMAIL")
	token = os.getenv("MARKETPLACE_API_KEY")
	
	if not email or not token:
		logging.error("MARKETPLACE_EMAIL or MARKETPLACE_API_KEY not found in .env file.")
		logging.error("Please create a .env file and add your marketplace credentials.")
		return
	
	headers = {
			"X-Marketplace-Email":        email,
			"X-Marketplace-Access-Token": token,
	}
	
	# --- Create Data Directories ---
	os.makedirs(CSV_DIR, exist_ok=True)
	os.makedirs(HTML_DIR, exist_ok=True)
	logging.info(f"Created data directories: {CSV_DIR}, {HTML_DIR}")
	
	# Generate timestamp for filenames
	now = datetime.now()
	timestamp = now.strftime("%Y-%m-%d_%H%M")
	
	# --- Fetch All Orders ---
	logging.info("Fetching all orders from marketplace…")
	try:
		resp = requests.get(f"{ORDERS_API_BASE_URL}/seller/orders",
		                    headers=headers,
		                    timeout=20)
		resp.raise_for_status()
		data = resp.json()
		orders_list: List[Dict[str, Any]] = data.get("data") or data.get("orders") or data
	except requests.exceptions.RequestException as exc:
		logging.error(f"API communication error: {exc}")
		if exc.response is not None:
			logging.error(f"Response content: {exc.response.text}")
		return
	except (ValueError, KeyError) as exc:
		logging.error(f"Error parsing API response: {exc}")
		return
	
	logging.info(f"Total orders returned from API: {len(orders_list)}")
	
	# --- Filter Orders Based on User Choice ---
	if filter_choice == "1":
		filtered_orders = [
				o for o in orders_list if o.get(FULFILLMENT_FIELD) != SHIPPED_VALUE
		]
		filter_description = "Not Shipped"
	elif filter_choice == "2":
		filtered_orders = [
				o for o in orders_list if o.get(FULFILLMENT_FIELD) == SHIPPED_VALUE
		]
		filter_description = "Shipped"
	else:  # choice == "3"
		filtered_orders = orders_list
		filter_description = "All"
	
	logging.info(f"Filtering for '{filter_description}' orders. Found {len(filtered_orders)} matching orders.")
	
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
		
		logging.info(f"Processing order {i + 1}/{total_to_process} (ID: {order_id})")
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
			
			image_uri = get_scryfall_image_uri(card_name, set_code, collector_number)
			
			# Handle image downloading if enabled
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
					"order_id":           order_id,
					"order_label":        details.get("label", "N/A"),
					"location":           location,
					"quantity":           item.get("quantity", 0),
					"name":               card_name,
					"set":                set_code,
					"number":             collector_number,
					"condition":          single.get("condition_id", "N/A"),
					"finish":             single.get("finish_id", "N/A"),
					"price":              item.get("price_cents", 0) / 100.0,
					"tcgplayer_sku":      product.get("tcgplayer_sku", "N/A"),
					"scryfall_image_uri": image_uri,
					"local_image_path":   local_image_path,
			})
	
	if not items:
		logging.info("No line items found in the processed orders — nothing exported.")
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
	
	# --- Determine output filename based on filter ---
	if filter_choice == "1":
		filter_description = "Not Shipped"
		output_filename = os.path.join(CSV_DIR, f"{timestamp}_orders_not-shipped.csv")
	elif filter_choice == "2":
		filter_description = "Shipped"
		output_filename = os.path.join(CSV_DIR, f"{timestamp}_orders_shipped.csv")
	else:
		filter_description = "All"
		output_filename = os.path.join(CSV_DIR, f"{timestamp}_orders_all.csv")
	
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
	
	# Sort by location and set for optimal picking workflow
	df.sort_values(by=['location_total_qty', 'location', 'set_total_qty', 'set', 'name'],
	               ascending=[False, True, False, True, True], inplace=True)
	
	df.drop(columns=['location_total_qty', 'set_total_qty'], inplace=True)
	
	try:
		df.to_csv(output_filename, index=False)
		logging.info(f"Successfully exported {len(items)} items to {output_filename}")
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
				print(f"📄 CSV file: {os.path.abspath(output_filename)}")
				print(f"🎨 HTML report: {os.path.abspath(html_filename)}")
			except Exception as exc:
				logging.warning(f"Could not open browser automatically: {exc}")
				print(f"\n📄 Files generated:")
				print(f"  • CSV: {os.path.abspath(output_filename)}")
				print(f"  • HTML: {os.path.abspath(html_filename)}")
				print(f"💡 Open {html_filename} in your browser to view the visual report!")
	else:
		print(f"\n📄 CSV file generated: {os.path.abspath(output_filename)}")
		print(f"💡 HTML report generation was disabled. Enable it in settings if you want visual reports.")
	
	logging.info("Order fulfillment process complete.")


if __name__ == "__main__":
	main()
