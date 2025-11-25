import sys
import requests


MANAPOOL_EMAIL = ""
MANAPOOL_ACCESS_TOKEN = ""

def get_unfulfilled_orders():
    """Get all unfulfilled orders."""
    if not MANAPOOL_EMAIL or not MANAPOOL_ACCESS_TOKEN:
        print("[X] Error: Set MANAPOOL_EMAIL and MANAPOOL_ACCESS_TOKEN at top of script")
        sys.exit(1)

    session = requests.Session()
    session.headers.update({
        'X-ManaPool-Email': MANAPOOL_EMAIL,
        'X-ManaPool-Access-Token': MANAPOOL_ACCESS_TOKEN,
    })

    orders = []
    offset = 0

    print("[*] Fetching unfulfilled orders...")
    while True:
        try:
            url = f"https://www.manapool.com/api/v1/seller/orders?limit=100&offset={offset}"
            resp = session.get(url, timeout=30)
            resp.raise_for_status()
            data = resp.json()


            batch = data.get('orders', [])

            if not batch:
                break

            for order in batch:
                status = order.get('latest_fulfillment_status')

                if status in (None, '', 'processing', 'unfulfilled'):
                    orders.append({
                        'id': order['id'],
                        'label': order['label'],
                        'total': f"${order['total_cents']/100:.2f}",
                        'status': status or 'unfulfilled'
                    })

            if len(batch) < 100:
                break

            offset += 100

        except Exception as e:
            print(f"[X] Error: {e}")
            sys.exit(1)

    return orders, session


def update_order_status(session, order_id, status, tracking_number=None):
    try:
        resp = session.put(
            f"https://www.manapool.com/api/v1/seller/orders/{order_id}/fulfillment",
            json={
                "status": status,
                "tracking_number": tracking_number,
                "tracking_company": None,
                "tracking_url": None,
                "in_transit_at": None,
                "estimated_delivery_at": None,
                "delivered_at": None,
            },
            timeout=30
        )
        return resp.status_code == 200
    except:
        return False


def main():
    orders, session = get_unfulfilled_orders()

    if not orders:
        print("[*] No unfulfilled orders found")
        return

    print(f"\n{'='*80}")
    print(f"Found {len(orders)} unfulfilled/processing orders")
    print(f"{'='*80}")

    for idx, order in enumerate(orders, 1):
        print(f"[{idx:3d}] {order['label']:15s} {order['total']:>10s}  {order['status']}")

    print(f"{'='*80}\n")

    print("What would you like to do?")
    print("  [1] Mark as processing")
    print("  [2] Mark as shipped")
    print("  [q] Quit")

    action = input("\nChoice (1/2/q): ").strip().lower()

    if action == 'q':
        return

    if action not in ('1', '2'):
        print("[X] Invalid choice")
        return

    new_status = 'processing' if action == '1' else 'shipped'

    print(f"\nEnter order numbers to mark as {new_status} (space-separated)")
    print("Examples:")
    print("  1 2 3        - Update orders 1, 2, and 3")
    print("  1-5          - Update orders 1 through 5")
    print("  1-5 10 15    - Update orders 1-5, 10, and 15")
    print("  all          - Update all orders")
    print("  q            - Cancel")

    choice = input(f"\nOrders to mark as {new_status}: ").strip().lower()

    if choice == 'q':
        return

    selected = set()
    if choice == 'all':
        selected = set(range(1, len(orders) + 1))
    else:
        for part in choice.split():
            if '-' in part:
                start, end = part.split('-')
                try:
                    selected.update(range(int(start), int(end) + 1))
                except:
                    print(f"[X] Invalid range: {part}")
                    return
            else:
                try:
                    selected.add(int(part))
                except:
                    print(f"[X] Invalid number: {part}")
                    return

    selected = [s for s in selected if 1 <= s <= len(orders)]
    if not selected:
        print("[X] No valid orders selected")
        return

    tracking = None
    if new_status == 'shipped':
        print(f"\n[*] Selected {len(selected)} orders")
        tracking = input("Enter tracking number for all (or press Enter to skip): ").strip()
        tracking = tracking if tracking else None

    print(f"\n[!] About to mark {len(selected)} orders as {new_status.upper()}")
    if tracking:
        print(f"[!] Tracking: {tracking}")
    confirm = input("Continue? (y/n): ").strip().lower()

    if confirm != 'y':
        print("[*] Cancelled")
        return

    print(f"\n[*] Updating {len(selected)} orders to {new_status}...")
    success = 0
    failed = 0

    for idx in sorted(selected):
        order = orders[idx - 1]
        if update_order_status(session, order['id'], new_status, tracking):
            print(f"[+] {order['label']}")
            success += 1
        else:
            print(f"[X] {order['label']} - FAILED")
            failed += 1

    print(f"\n{'='*80}")
    print(f"[*] Done: {success} updated, {failed} failed")
    print(f"{'='*80}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[*] Cancelled")
        sys.exit(0)
