import sys
import requests
import logging
import pandas as pd
import time
from typing import Any, Dict, List, Optional

# Configuration

LOG_FILE = "order_retrieval_log.txt"
# The output CSV filename will now be generated dynamically based on the filter.
BASE_OUTPUT_CSV_NAME = "fulfillment_list"
EMAIL_FILE = "email.txt"  # ManaPool user email
API_KEY_FILE = "api-key.txt"  # ManaPool API token

MANAPOOL_API_BASE_URL = "https://manapool.com/api/v1"
SCRYFALL_API_BASE_URL = "https://api.scryfall.com"

FULFILLMENT_FIELD = "latest_fulfillment_status"
SHIPPED_VALUE = "shipped"  # Skip orders with this value


def read_credential_from_file(filename: str) -> Optional[str]:
    """Read and return the first line of *filename*, stripped of whitespace."""
    try:
        with open(filename, "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except FileNotFoundError:
        logging.error(
            f"Credential file not found: {filename} — create it in the script directory."
        )
        return None


def get_scryfall_image_uri(card_name: str, set_code: str,
                           collector_number: str) -> str:
    """Return a normal‑sized image URL from Scryfall (or 'N/A' if not found)."""
    time.sleep(0.1)

    # First try exact set / collector lookup
    try:
        url = f"{SCRYFALL_API_BASE_URL}/cards/{set_code.lower()}/{collector_number}"
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return r.json().get("image_uris", {}).get("normal", "N/A")
    except requests.exceptions.RequestException:
        pass

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


def get_user_filter_choice() -> str:
    """Prompt the user to select an order filter and return their choice."""
    print("\n" + "=" * 40)
    print("  ManaPool Order Retrieval")
    print("=" * 40)
    while True:
        print("\nPlease select which orders to retrieve:")
        print("  1: Not Shipped (Default)")
        print("  2: Shipped Only")
        print("  3: All Orders")
        choice = input("Enter your choice (1-3): ").strip()
        if choice in ["1", "2", "3"]:
            return choice
        print("\nInvalid choice. Please enter 1, 2, or 3.")


def main() -> None:
    # Logging setup
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        filename=LOG_FILE,
        filemode="w",
    )
    logging.getLogger().addHandler(logging.StreamHandler(sys.stdout))

    # --- Get User Input ---
    filter_choice = get_user_filter_choice()

    # Credentials
    email = read_credential_from_file(EMAIL_FILE)
    token = read_credential_from_file(API_KEY_FILE)
    if not email or not token:
        logging.error("Missing credentials — aborting.")
        return

    headers = {
        "X-ManaPool-Email": email,
        "X-ManaPool-Access-Token": token,
    }

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
        return
    except (ValueError, KeyError) as exc:
        logging.error(f"Error parsing API response: {exc}")
        return

    logging.info(f"Total orders returned from API: {len(orders_list)}")

    # --- Filter Orders Based on User Choice ---
    if filter_choice == "1":
        filtered_orders = [o for o in orders_list if o.get(FULFILLMENT_FIELD) != SHIPPED_VALUE]
        filter_description = "Not Shipped"
        output_filename = f"{BASE_OUTPUT_CSV_NAME}_not_shipped.csv"
    elif filter_choice == "2":
        filtered_orders = [o for o in orders_list if o.get(FULFILLMENT_FIELD) == SHIPPED_VALUE]
        filter_description = "Shipped"
        output_filename = f"{BASE_OUTPUT_CSV_NAME}_shipped.csv"
    else:  # choice == "3"
        filtered_orders = orders_list
        filter_description = "All"
        output_filename = f"{BASE_OUTPUT_CSV_NAME}_all.csv"

    logging.info(f"Filtering for '{filter_description}' orders. Found {len(filtered_orders)} matching orders.")

    if not filtered_orders:
        logging.info("No orders match the selected filter — exiting.")
        return

    items: List[Dict[str, Any]] = []
    total_to_process = len(filtered_orders)
    logging.info(f"Fetching details for {total_to_process} orders...")

    for i, summary in enumerate(filtered_orders):
        order_id = summary.get("id")
        if not order_id:
            continue

        logging.info(f"Processing order {i+1}/{total_to_process} (ID: {order_id})")
        details = get_order_details(order_id, headers)
        if not details:
            continue

        for item in details.get("items", []):
            product = item.get("product", {})
            single = product.get("single", {})

            card_name = single.get("name", "N/A")
            set_code = single.get("set", "N/A")
            collector_number = single.get("number", "N/A")

            image_uri = get_scryfall_image_uri(card_name, set_code,
                                               collector_number)

            items.append({
                "order_id": order_id,
                "order_label": details.get("label", "N/A"),
                "quantity": item.get("quantity", 0),
                "name": card_name,
                "set": set_code,
                "number": collector_number,
                "condition": single.get("condition_id", "N/A"),
                "finish": single.get("finish_id", "N/A"),
                "price": item.get("price_cents", 0) / 100.0,
                "tcgplayer_sku": product.get("tcgplayer_sku", "N/A"),
                "scryfall_image_uri": image_uri,
            })

    if not items:
        logging.info("No line items found in the processed orders — nothing exported.")
        return

    df = pd.DataFrame(items)
    try:
        df.to_csv(output_filename, index=False)
        logging.info(f"Successfully exported {len(items)} items to {output_filename}")
    except Exception as exc:
        logging.error(f"Failed to write CSV file: {exc}")

    logging.info("Order‑retrieval process complete.")


if __name__ == "__main__":
    main()
