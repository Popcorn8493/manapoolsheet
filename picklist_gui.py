

import csv
import json
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import requests
from io import BytesIO
import threading
import time

class PicklistGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MTG Picklist Manager")
        
        self.setup_window_size_and_position()
        
        self.root.minsize(1200, 800)
        
        self.setup_styles()
        
        self.cards = []
        self.filtered_cards = []
        self.locations = {}
        self.grabbed_cards = set()
        self.image_cache = {}
        self.price_cache = {}
        self.current_filter = ""
        self.current_location_filter = "All"
        self.current_sort = "order"
        
        self.load_locations()
        
        self.create_widgets()
        
        self.root.bind('<Configure>', self.on_window_resize)
    
    def setup_window_size_and_position(self):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = int(screen_width * 0.7)
        window_height = int(screen_height * 0.7)
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
    
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        self.colors = {
            'primary': '#2E86AB',
            'secondary': '#A23B72',
            'success': '#F18F01',
            'warning': '#C73E1D',
            'background': '#F8F9FA',
            'surface': '#FFFFFF',
            'text': '#212529',
            'text_secondary': '#6C757D'
        }
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'), foreground=self.colors['text'])
        style.configure('Title.TLabel', font=('Segoe UI', 10, 'bold'), foreground=self.colors['primary'])
        style.configure('Card.TLabel', font=('Segoe UI', 9), foreground=self.colors['text'])
        style.configure('Small.TLabel', font=('Segoe UI', 8), foreground=self.colors['text_secondary'])
        style.configure('Success.TLabel', font=('Segoe UI', 9, 'bold'), foreground=self.colors['success'])
        style.configure('Primary.TButton', font=('Segoe UI', 9, 'bold'))
        style.configure('Secondary.TButton', font=('Segoe UI', 9))
        style.configure('Card.TFrame', relief='solid', borderwidth=1)
        try:
            style.configure('Location.TLabelFrame', 
                           font=('Segoe UI', 10, 'bold'), 
                           foreground=self.colors['primary'],
                           relief='solid',
                           borderwidth=1)
            style.configure('Location.TLabelFrame.Label',
                           font=('Segoe UI', 10, 'bold'),
                           foreground=self.colors['primary'])
        except tk.TclError:
            pass
        
    def load_locations(self):
        try:
            if os.path.exists("locations.json"):
                with open("locations.json", "r", encoding="utf-8") as f:
                    self.locations = json.load(f)
            else:
                self.locations = {}
                with open("locations.json", "w", encoding="utf-8") as f:
                    json.dump(self.locations, f, indent=2)
        except Exception as e:
            print(f"Could not load locations: {e}")
            self.locations = {}
    
    def create_widgets(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_container = ttk.Frame(self.root, padding="10")
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(2, weight=1)
        self.create_header_section(main_container)
        self.create_filter_section(main_container)
        self.create_content_area(main_container)
        self.create_status_bar(main_container)
    
    def create_header_section(self, parent):
        header_frame = ttk.Frame(parent)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        header_frame.columnconfigure(1, weight=1)
        file_section = ttk.Frame(header_frame)
        file_section.grid(row=0, column=0, sticky=tk.W)
        ttk.Button(file_section, text="Load CSV", command=self.load_picklist, 
                  style='Primary.TButton').pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(file_section, text="Manage Locations", command=self.manage_locations, 
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 10))
        self.file_label = ttk.Label(file_section, text="No file loaded", style='Small.TLabel')
        self.file_label.pack(side=tk.LEFT)
        stats_section = ttk.Frame(header_frame)
        stats_section.grid(row=0, column=1, sticky=tk.E)
        self.stats_label = ttk.Label(stats_section, text="", style='Small.TLabel')
        self.stats_label.pack(side=tk.RIGHT)
        self.progress_var = tk.StringVar(value="Ready")
        ttk.Label(stats_section, textvariable=self.progress_var, style='Small.TLabel').pack(side=tk.RIGHT, padx=(0, 10))
    
    def create_filter_section(self, parent):
        filter_frame = ttk.LabelFrame(parent, text="Filters & Search", padding="10")
        filter_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        filter_frame.columnconfigure(1, weight=1)
        ttk.Label(filter_frame, text="Search:", style='Card.TLabel').grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.apply_filters)
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_var, width=30)
        search_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 20))
        ttk.Label(filter_frame, text="Location:", style='Card.TLabel').grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        self.location_var = tk.StringVar(value="All")
        self.location_var.trace('w', self.apply_filters)
        location_combo = ttk.Combobox(filter_frame, textvariable=self.location_var, state="readonly", width=15)
        location_combo.grid(row=0, column=3, sticky=tk.W)
        filter_buttons = ttk.Frame(filter_frame)
        filter_buttons.grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))
        ttk.Button(filter_buttons, text="All", command=lambda: self.set_filter("all"), 
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(filter_buttons, text="Remaining", command=lambda: self.set_filter("remaining"), 
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(filter_buttons, text="Grabbed", command=lambda: self.set_filter("grabbed"), 
                  style='Secondary.TButton').pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(filter_buttons, text="Clear Filters", command=self.clear_filters, 
                  style='Secondary.TButton').pack(side=tk.LEFT)
        sort_frame = ttk.Frame(filter_frame)
        sort_frame.grid(row=2, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))
        ttk.Label(sort_frame, text="Sort:", style='Card.TLabel').pack(side=tk.LEFT, padx=(0, 5))
        self.sort_var = tk.StringVar(value="order")
        self.sort_var.trace('w', self.apply_filters)
        sort_combo = ttk.Combobox(sort_frame, textvariable=self.sort_var, state="readonly", width=12)
        sort_combo['values'] = [
            "Order",
            "Card Name", 
            "Location",
            "Set Code",
            "Rarity",
            "Condition",
            "Price"
        ]
        sort_combo.pack(side=tk.LEFT, padx=(0, 5))
        self.sort_direction = tk.StringVar(value="asc")
        self.direction_button = ttk.Button(sort_frame, text="↑", width=3, 
                                         command=self.toggle_sort_direction)
        self.direction_button.pack(side=tk.LEFT)
    
    def create_content_area(self, parent):
        self.notebook = ttk.Notebook(parent)
        self.notebook.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.cards_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.cards_frame, text="Cards")
        self.create_cards_display()
        self.summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.summary_frame, text="Summary")
        self.create_summary_display()
    
    def create_status_bar(self, parent):
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status_var, style='Small.TLabel').pack(side=tk.LEFT)
        self.progress_bar = ttk.Progressbar(status_frame, mode='indeterminate', length=200)
        self.progress_bar.pack(side=tk.RIGHT)
        
    def create_cards_display(self):
        self.cards_container = ttk.Frame(self.cards_frame, padding="10")
        self.cards_container.pack(fill=tk.BOTH, expand=True)
        self.cards_canvas = tk.Canvas(self.cards_container, bg=self.colors['background'])
        self.cards_scrollbar = ttk.Scrollbar(self.cards_container, orient="vertical", command=self.cards_canvas.yview)
        self.cards_scrollable_frame = ttk.Frame(self.cards_canvas)
        self.cards_scrollable_frame.bind(
            "<Configure>",
            lambda e: self.cards_canvas.configure(scrollregion=self.cards_canvas.bbox("all"))
        )
        self.cards_canvas.create_window((0, 0), window=self.cards_scrollable_frame, anchor="nw")
        self.cards_canvas.configure(yscrollcommand=self.cards_scrollbar.set)
        self.cards_canvas.pack(side="left", fill="both", expand=True)
        self.cards_scrollbar.pack(side="right", fill="y")
        def _on_mousewheel(event):
            self.cards_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.cards_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.show_empty_state()
    
    def show_empty_state(self):
        for widget in self.cards_scrollable_frame.winfo_children():
            widget.destroy()
        empty_frame = ttk.Frame(self.cards_scrollable_frame)
        empty_frame.pack(expand=True, fill=tk.BOTH, pady=50)
        ttk.Label(empty_frame, text="No cards loaded", style='Header.TLabel').pack(pady=20)
        ttk.Label(empty_frame, text="Click 'Load CSV' to import a picklist", style='Small.TLabel').pack()
    
    def apply_filters(self, *args):
        if not self.cards:
            return
        search_text = self.search_var.get().lower()
        location_filter = self.location_var.get()
        self.filtered_cards = []
        for card in self.cards:
            if search_text:
                if (search_text not in card['name'].lower() and 
                    search_text not in card['set'].lower() and
                    search_text not in card['set_code'].lower()):
                    continue
            if location_filter != "All" and card['location'] != location_filter:
                continue
            self.filtered_cards.append(card)
        self.sort_cards()
        self.update_location_dropdown()
        self.display_cards()
        self.update_stats()
    
    def sort_cards(self):
        if not self.filtered_cards:
            return
        sort_option = self.sort_var.get()
        direction = self.sort_direction.get()
        reverse = (direction == "desc")
        if sort_option == "Order":
            def order_key(card):
                order = card['order']
                try:
                    return int(order)
                except (ValueError, TypeError):
                    return str(order)
            self.filtered_cards.sort(key=order_key, reverse=reverse)
        elif sort_option == "Card Name":
            self.filtered_cards.sort(key=lambda x: x['name'].lower(), reverse=reverse)
        elif sort_option == "Location":
            self.filtered_cards.sort(key=lambda x: x['location'].lower(), reverse=reverse)
        elif sort_option == "Set Code":
            self.filtered_cards.sort(key=lambda x: x['set_code'].lower(), reverse=reverse)
        elif sort_option == "Rarity":
            self.filtered_cards.sort(key=lambda x: x['rarity'].lower(), reverse=reverse)
        elif sort_option == "Condition":
            self.filtered_cards.sort(key=lambda x: x['condition'].lower(), reverse=reverse)
        elif sort_option == "Price":
            def price_key(card):
                price = card.get('price')
                if price is None:
                    return float('inf') if reverse else float('-inf')
                try:
                    return float(price)
                except (ValueError, TypeError):
                    return float('inf') if reverse else float('-inf')
            self.filtered_cards.sort(key=price_key, reverse=reverse)
    
    def set_filter(self, filter_type):
        if filter_type == "all":
            self.search_var.set("")
            self.location_var.set("All")
        elif filter_type == "remaining":
            self.search_var.set("")
            self.location_var.set("All")
            self.current_filter = "remaining"
        elif filter_type == "grabbed":
            self.search_var.set("")
            self.location_var.set("All")
            self.current_filter = "grabbed"
        self.apply_filters()
    
    def toggle_sort_direction(self):
        current = self.sort_direction.get()
        new_direction = "desc" if current == "asc" else "asc"
        self.sort_direction.set(new_direction)
        self.direction_button.config(text="↓" if new_direction == "desc" else "↑")
        self.apply_filters()
    
    def clear_filters(self):
        self.search_var.set("")
        self.location_var.set("All")
        self.sort_var.set("Order")
        self.sort_direction.set("asc")
        self.direction_button.config(text="↑")
        self.current_filter = ""
        self.apply_filters()
    
    def update_location_dropdown(self):
        locations = set(card['location'] for card in self.filtered_cards)
        locations = ["All"] + sorted(locations)
        location_combo = None
        for child in self.cards_frame.winfo_children():
            if isinstance(child, ttk.LabelFrame):
                for widget in child.winfo_children():
                    if isinstance(widget, ttk.Combobox):
                        location_combo = widget
                        break
        if location_combo:
            location_combo['values'] = locations
    
    def update_stats(self):
        if not self.cards:
            self.stats_label.config(text="")
            return
        total = len(self.cards)
        grabbed = sum(1 for card in self.cards if card['grabbed'])
        remaining = total - grabbed
        stats_text = f"Total: {total} | Grabbed: {grabbed} | Remaining: {remaining}"
        if self.filtered_cards:
            filtered_count = len(self.filtered_cards)
            stats_text += f" | Showing: {filtered_count}"
        self.stats_label.config(text=stats_text)
        
    def create_summary_display(self):
        self.summary_text = tk.Text(self.summary_frame, wrap=tk.WORD, height=20)
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        buttons_frame = ttk.Frame(self.summary_frame)
        buttons_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Button(buttons_frame, text="Refresh Summary", command=self.update_summary).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame, text="Save Progress", command=self.save_progress).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_frame, text="Load Progress", command=self.load_progress).pack(side=tk.LEFT)
        
    def load_picklist(self):
        
        filename = filedialog.askopenfilename(
            title="Select Picklist CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not filename:
            return
        
        
        self.progress_bar.start()
        self.progress_var.set("Loading...")
        self.status_var.set("Loading CSV file...")
        
        try:
            self.cards = []
            self.grabbed_cards.clear()
            self.filtered_cards = []
            
            with open(filename, 'r', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    
                    set_code = row.get('Set Code', '')
                    location = self.get_location_for_set(set_code)
                    
                    card = {
                        'order': row.get('Order', ''),
                        'name': row.get('Card Name', ''),
                        'set': row.get('Set', ''),
                        'set_code': set_code,
                        'collector_number': row.get('Collector Number', ''),
                        'quantity': int(row.get('Quantity', 1)),
                        'condition': row.get('Condition', ''),
                        'language': row.get('Language', ''),
                        'finish': row.get('Finish', ''),
                        'rarity': row.get('Rarity', ''),
                        'location': location,
                        'grabbed': False,
                        'price': None,  
                        'price_loading': False
                    }
                    self.cards.append(card)
            

            self.file_label.config(text=f"📁 {os.path.basename(filename)} ({len(self.cards)} cards)")
            self.progress_var.set("Ready")
            self.status_var.set(f"Loaded {len(self.cards)} cards successfully")
            

            self.apply_filters()
            self.update_summary()
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not load CSV file: {e}")
            self.status_var.set("Error loading file")
        finally:
            self.progress_bar.stop()
    
    def get_location_for_set(self, set_code):
        
        if not set_code:
            return "Unknown"
        return self.locations.get(set_code.upper(), "Unassigned")
    
    def display_cards(self):
        
        for widget in self.cards_scrollable_frame.winfo_children():
            widget.destroy()
        
        if not self.cards:
            self.show_empty_state()
            return
        
        cards_to_display = self.filtered_cards if self.filtered_cards else self.cards
        
        if self.current_filter == "remaining":
            cards_to_display = [card for card in cards_to_display if not card['grabbed']]
        elif self.current_filter == "grabbed":
            cards_to_display = [card for card in cards_to_display if card['grabbed']]
        
        if not cards_to_display:
            self.show_no_results_state()
            return
        
        sort_option = self.sort_var.get()

        should_group_by_location = (sort_option == "Location")
        
        if should_group_by_location:

            cards_by_location = {}
            for card in cards_to_display:
                location = card['location']
                if location not in cards_by_location:
                    cards_by_location[location] = []
                cards_by_location[location].append(card)
            
            
            for location, location_cards in sorted(cards_by_location.items()):
                self.create_location_section(location, location_cards)
        else:
            
            self.create_flat_card_display(cards_to_display)
    
    def show_no_results_state(self):
        
        for widget in self.cards_scrollable_frame.winfo_children():
            widget.destroy()
        
        empty_frame = ttk.Frame(self.cards_scrollable_frame)
        empty_frame.pack(expand=True, fill=tk.BOTH, pady=50)
        
        ttk.Label(empty_frame, text="🔍 No cards match your filters", style='Header.TLabel').pack(pady=20)
        ttk.Label(empty_frame, text="Try adjusting your search or filter criteria", style='Small.TLabel').pack()
        ttk.Button(empty_frame, text="Clear Filters", command=self.clear_filters, 
                  style='Primary.TButton').pack(pady=10)
    
    def create_location_section(self, location, cards):
        
        grabbed_count = sum(1 for card in cards if card['grabbed'])
        total_count = len(cards)
        progress_percent = (grabbed_count / total_count * 100) if total_count > 0 else 0
        
        location_frame = ttk.LabelFrame(
            self.cards_scrollable_frame, 
            text=f"{location} ({grabbed_count}/{total_count} - {progress_percent:.1f}%)",
            padding="10"
        )
        
        location_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.create_card_grid(location_frame, cards)
    
    def create_flat_card_display(self, cards):
        
        main_frame = ttk.Frame(self.cards_scrollable_frame, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        sort_option = self.sort_var.get()
        direction = "↑" if self.sort_direction.get() == "asc" else "↓"
        
        if sort_option == "Order":
            header_text = f"📋 Cards in original picklist order {direction}"
        else:
            header_text = f"📋 Cards sorted by: {sort_option} {direction}"
        
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        sort_indicator = ttk.Frame(header_frame, style='Card.TFrame')
        sort_indicator.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(sort_indicator, text=header_text, style='Header.TLabel').pack(side=tk.LEFT, padx=10, pady=5)
        
        total_count = len(cards)
        grabbed_count = sum(1 for card in cards if card['grabbed'])
        count_text = f"({grabbed_count}/{total_count} grabbed)"
        ttk.Label(sort_indicator, text=count_text, style='Small.TLabel').pack(side=tk.RIGHT, padx=10, pady=5)
        
        self.create_card_grid(main_frame, cards)
    
    def create_card_grid(self, parent, cards):
        
        parent.update_idletasks()
        
        card_width = 180  
        
        try:
            main_width = self.root.winfo_width()
            available_width = main_width - 250
        except:
            available_width = 1000  
        
        cards_per_row = max(1, available_width // (card_width + 10))
        cards_per_row = min(cards_per_row, 6)  
        cards_per_row = max(cards_per_row, 2)  
        
        current_row_frame = None
        for i, card in enumerate(cards):
            row = i // cards_per_row
            col = i % cards_per_row
            
            if col == 0:  
                current_row_frame = ttk.Frame(parent)
                current_row_frame.pack(fill=tk.X, pady=2)
                
                for c in range(cards_per_row):
                    current_row_frame.columnconfigure(c, weight=1)
            
            self.create_modern_card_widget(current_row_frame, card, col)
    
    def create_modern_card_widget(self, parent, card, column):
        
        
        card_frame = ttk.Frame(parent, style='Card.TFrame', padding="8")
        card_frame.grid(row=0, column=column, sticky=(tk.W, tk.E, tk.N, tk.S), padx=2, pady=2)
        
        
        parent.columnconfigure(column, weight=1)
        
        
        header_frame = ttk.Frame(card_frame)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        
        
        status_color = self.colors['success'] if card['grabbed'] else self.colors['text_secondary']
        status_canvas = tk.Canvas(header_frame, width=12, height=12, highlightthickness=0)
        status_canvas.pack(side=tk.LEFT)
        status_canvas.create_oval(2, 2, 10, 10, fill=status_color, outline="")
        
        
        order_short = card['order'].split('-')[0] if '-' in card['order'] else card['order'][:8]
        ttk.Label(header_frame, text=f"#{order_short}", style='Small.TLabel').pack(side=tk.LEFT, padx=(4, 0))

        image_frame = ttk.Frame(card_frame)
        image_frame.pack(fill=tk.X, pady=(0, 8))

        image_label = ttk.Label(image_frame, text="🃏", font=('Arial', 24))
        image_label.pack()
        
        
        def load_with_delay():
            time.sleep(0.1)  
            
            try:
                if image_label.winfo_exists():
                    self.load_card_image(card, image_label)
            except tk.TclError:
                
                pass
        
        threading.Thread(target=load_with_delay, daemon=True).start()
        
        
        details_frame = ttk.Frame(card_frame)
        details_frame.pack(fill=tk.BOTH, expand=True)
        
        
        name_label = ttk.Label(details_frame, text=card['name'], style='Card.TLabel', wraplength=180)
        name_label.pack(anchor=tk.W, pady=(0, 4))
        
        set_info = f"{card['set_code']}"
        set_label = ttk.Label(details_frame, text=set_info, style='Small.TLabel')
        set_label.pack(anchor=tk.W, pady=(0, 2))
        qty_cond = f"{card['quantity']}x • {card['condition']}"
        qty_label = ttk.Label(details_frame, text=qty_cond, style='Small.TLabel')
        qty_label.pack(anchor=tk.W, pady=(0, 4))
        
        
        details = []
        if card['rarity'] and card['rarity'] != 'N/A':
            details.append(card['rarity'])
        if card['language'] and card['language'] != 'N/A':
            details.append(card['language'])
        if card['finish'] and card['finish'] != 'N/A':
            
            finish = card['finish']
            if 'FOIL' in finish.upper():
                finish = finish.replace('*', '').strip()
                if finish.upper() == 'FOIL':
                    finish = '★ FOIL'
                else:
                    finish = f"★ {finish}"
            details.append(finish)
        
        if details:
            details_text = " • ".join(details)
            details_label = ttk.Label(details_frame, text=details_text, style='Small.TLabel')
            details_label.pack(anchor=tk.W, pady=(0, 4))
        
        
        price_label = ttk.Label(details_frame, text="💰 Loading price...", style='Small.TLabel')
        price_label.pack(anchor=tk.W, pady=(0, 8))
        
        
        card['price_label'] = price_label
        
        
        def load_price_with_delay():
            time.sleep(0.2)  
            try:
                if price_label.winfo_exists():
                    price = self.fetch_card_price(card)
                    card['price'] = price  
                    if price is not None:
                        price_text = f"💰 ${price:.2f}"
                        self.root.after(0, lambda: self.update_price_label(price_label, price_text))
                    else:
                        self.root.after(0, lambda: self.update_price_label(price_label, "💰 Price N/A"))
            except tk.TclError:
                
                pass
        
        threading.Thread(target=load_price_with_delay, daemon=True).start()
        
        
        action_frame = ttk.Frame(details_frame)
        action_frame.pack(fill=tk.X, pady=(0, 4))
        
        
        grabbed_var = tk.BooleanVar(value=card['grabbed'])
        grabbed_check = ttk.Checkbutton(
            action_frame, 
            text="✓ Grabbed" if card['grabbed'] else "Mark as Grabbed",
            variable=grabbed_var,
            command=lambda: self.toggle_grabbed(card, grabbed_var.get()),
            style='Success.TLabel' if card['grabbed'] else 'Card.TLabel'
        )
        grabbed_check.pack(side=tk.LEFT)
        
        
        if card['grabbed']:
            card_frame.configure(relief=tk.SUNKEN, borderwidth=2)
        else:
            card_frame.configure(relief=tk.RAISED, borderwidth=1)
        
        
        card['widget'] = card_frame
        card['grabbed_var'] = grabbed_var
        card['image_label'] = image_label
    
    def fetch_card_price(self, card):
        
        try:
            
            cache_key = f"{card['set_code']}_{card['collector_number']}_{card['name']}"
            if cache_key in self.price_cache:
                return self.price_cache[cache_key]
            
            
            price = None
            
            
            set_code = card['set_code'].lower()
            collector_number = card['collector_number']
            
            if set_code and collector_number and set_code != 'n/a' and collector_number != 'n/a':
                try:
                    url = f"https://api.scryfall.com/cards/{set_code}/{collector_number}"
                    response = requests.get(url, timeout=10)
                    if response.status_code == 200:
                        data = response.json()
                        price = self.extract_price_from_data(data)
                except:
                    pass
            
            
            if not price:
                card_name = card['name']
                if card_name and card_name != 'N/A':
                    try:
                        url = f"https://api.scryfall.com/cards/named?fuzzy={card_name}"
                        response = requests.get(url, timeout=10)
                        if response.status_code == 200:
                            data = response.json()
                            price = self.extract_price_from_data(data)
                    except:
                        pass
            
            
            self.price_cache[cache_key] = price
            return price
            
        except Exception as e:
            print(f"Error fetching price for {card.get('name', 'Unknown')}: {e}")
            return None
    
    def extract_price_from_data(self, data):
        
        try:
            
            prices = data.get('prices', {})
            
            
            for price_type in ['usd', 'usd_foil', 'eur', 'eur_foil']:
                if price_type in prices and prices[price_type]:
                    return float(prices[price_type])
            
            return None
        except (ValueError, TypeError, KeyError):
            return None
    
    def load_card_image(self, card, image_label):
        
        try:
            
            cache_key = f"{card['set_code']}_{card['collector_number']}_{card['name']}"
            if cache_key in self.image_cache:
                self.root.after(0, lambda: self.update_image_label(image_label, self.image_cache[cache_key]))
                return
            
            
            image_url = None
            
            
            set_code = card['set_code'].lower()
            collector_number = card['collector_number']
            
            if set_code and collector_number and set_code != 'n/a' and collector_number != 'n/a':
                try:
                    url = f"https://api.scryfall.com/cards/{set_code}/{collector_number}"
                    response = requests.get(url, timeout=10)
                    if response.status_code == 200:
                        data = response.json()
                        image_url = data.get('image_uris', {}).get('normal', '')
                except:
                    pass
            
            
            if not image_url:
                card_name = card['name']
                if card_name and card_name != 'N/A':
                    try:
                        url = f"https://api.scryfall.com/cards/named?fuzzy={card_name}"
                        response = requests.get(url, timeout=10)
                        if response.status_code == 200:
                            data = response.json()
                            image_url = data.get('image_uris', {}).get('normal', '')
                    except:
                        pass
            
            
            if not image_url and set_code and set_code != 'n/a':
                card_name = card['name']
                if card_name and card_name != 'N/A':
                    try:
                        url = f"https://api.scryfall.com/cards/search?q=name:{card_name}+set:{set_code}"
                        response = requests.get(url, timeout=10)
                        if response.status_code == 200:
                            data = response.json()
                            cards = data.get('data', [])
                            if cards:
                                image_url = cards[0].get('image_uris', {}).get('normal', '')
                    except:
                        pass
            
            
            if image_url:
                img_response = requests.get(image_url, timeout=10)
                if img_response.status_code == 200:
                    image = Image.open(BytesIO(img_response.content))
                    
                    image.thumbnail((120, 168), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(image)
                    
                    
                    self.image_cache[cache_key] = photo
                    
                    
                    self.root.after(0, lambda: self.update_image_label(image_label, photo))
                    return
            
            
            self.root.after(0, lambda: image_label.config(text="🃏", image="", font=('Arial', 24)))
            
        except Exception as e:
            print(f"Error loading image for {card.get('name', 'Unknown')}: {e}")
            self.root.after(0, lambda: image_label.config(text="❌", image="", font=('Arial', 24)))
    
    def update_image_label(self, label, photo):
        
        try:
            
            if label.winfo_exists():
                label.config(image=photo, text="")
                label.image = photo  
        except tk.TclError:
            
            pass
    
    def update_price_label(self, label, price_text):
        
        try:
            
            if label.winfo_exists():
                label.config(text=price_text)
        except tk.TclError:
            
            pass
    
    def toggle_grabbed(self, card, grabbed):
        
        card['grabbed'] = grabbed
        if grabbed:
            self.grabbed_cards.add(card['order'] + card['name'])
        else:
            self.grabbed_cards.discard(card['order'] + card['name'])
        
        
        if grabbed:
            card['widget'].configure(relief=tk.SUNKEN, borderwidth=2)
            
            for widget in card['widget'].winfo_children():
                if isinstance(widget, ttk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Checkbutton):
                            child.configure(text="✓ Grabbed", style='Success.TLabel')
        else:
            card['widget'].configure(relief=tk.RAISED, borderwidth=1)
            
            for widget in card['widget'].winfo_children():
                if isinstance(widget, ttk.Frame):
                    for child in widget.winfo_children():
                        if isinstance(child, ttk.Checkbutton):
                            child.configure(text="Mark as Grabbed", style='Card.TLabel')
        
        
        self.update_stats()
        
        
        self.status_var.set(f"Card '{card['name']}' {'marked as grabbed' if grabbed else 'unmarked'}")
    
    def manage_locations(self):
        
        
        location_window = tk.Toplevel(self.root)
        location_window.title("Manage Locations")
        location_window.geometry("600x400")
        location_window.transient(self.root)
        location_window.grab_set()
        
        
        main_frame = ttk.Frame(location_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        
        ttk.Label(main_frame, text="Set Location for Each Set Code", style='Header.TLabel').pack(pady=(0, 20))
        ttk.Label(main_frame, text="Configure where each set is stored in your collection", style='Small.TLabel').pack(pady=(0, 20))
        
        
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        
        set_codes = set()
        if self.cards:
            set_codes = set(card['set_code'] for card in self.cards if card['set_code'] and card['set_code'] != 'N/A')
        
        
        location_vars = {}
        for i, set_code in enumerate(sorted(set_codes)):
            entry_frame = ttk.Frame(scrollable_frame)
            entry_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(entry_frame, text=f"{set_code}:", width=15, anchor=tk.W).pack(side=tk.LEFT)
            
            location_var = tk.StringVar(value=self.locations.get(set_code, "Unassigned"))
            location_entry = ttk.Entry(entry_frame, textvariable=location_var, width=30)
            location_entry.pack(side=tk.LEFT, padx=(10, 0))
            
            location_vars[set_code] = location_var
        
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        def save_locations():
            
            for set_code, var in location_vars.items():
                self.locations[set_code] = var.get()
            
            
            try:
                with open("locations.json", "w", encoding="utf-8") as f:
                    json.dump(self.locations, f, indent=2)
                
                
                for card in self.cards:
                    if card['set_code'] in self.locations:
                        card['location'] = self.locations[card['set_code']]
                
                
                self.apply_filters()
                messagebox.showinfo("Success", "Locations saved successfully!")
                location_window.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Could not save locations: {e}")
        
        ttk.Button(button_frame, text="Save Locations", command=save_locations, 
                  style='Primary.TButton').pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Cancel", command=location_window.destroy, 
                  style='Secondary.TButton').pack(side=tk.LEFT)
    
    def on_window_resize(self, event):
        
        if event.widget == self.root:
            
            if hasattr(self, '_resize_timer'):
                self.root.after_cancel(self._resize_timer)
            
            self._resize_timer = self.root.after(300, self.refresh_card_layout)
    
    def refresh_card_layout(self):
        
        if self.cards and hasattr(self, 'cards_scrollable_frame'):
            
            if self.cards_scrollable_frame.winfo_children():
                self.display_cards()
    
    def update_summary(self):
        
        if not self.cards:
            self.summary_text.delete(1.0, tk.END)
            self.summary_text.insert(tk.END, "No cards loaded")
            return
        
        
        total_cards = len(self.cards)
        grabbed_count = sum(1 for card in self.cards if card['grabbed'])
        remaining_count = total_cards - grabbed_count
        
        
        location_stats = {}
        for card in self.cards:
            location = card['location']
            if location not in location_stats:
                location_stats[location] = {'total': 0, 'grabbed': 0}
            location_stats[location]['total'] += 1
            if card['grabbed']:
                location_stats[location]['grabbed'] += 1
        
        
        summary = f"PICKLIST SUMMARY\n"
        summary += f"{'='*50}\n\n"
        summary += f"Total Cards: {total_cards}\n"
        summary += f"Grabbed: {grabbed_count}\n"
        summary += f"Remaining: {remaining_count}\n\n"
        
        summary += f"BY LOCATION:\n"
        summary += f"{'-'*30}\n"
        for location, stats in sorted(location_stats.items()):
            progress = (stats['grabbed'] / stats['total']) * 100 if stats['total'] > 0 else 0
            summary += f"{location}: {stats['grabbed']}/{stats['total']} ({progress:.1f}%)\n"
        
        summary += f"\nREMAINING CARDS:\n"
        summary += f"{'-'*30}\n"
        for card in self.cards:
            if not card['grabbed']:
                summary += f"• {card['name']} [{card['set_code']}] - {card['location']}\n"
        
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.insert(tk.END, summary)
    
    def save_progress(self):
        
        if not self.cards:
            messagebox.showwarning("Warning", "No cards loaded to save")
            return
        
        filename = filedialog.asksaveasfilename(
            title="Save Progress",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                progress_data = {
                    'grabbed_cards': list(self.grabbed_cards),
                    'card_states': {f"{card['order']}{card['name']}": card['grabbed'] for card in self.cards}
                }
                
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(progress_data, f, indent=2)
                
                messagebox.showinfo("Success", f"Progress saved to {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Could not save progress: {e}")
    
    def load_progress(self):
        
        filename = filedialog.askopenfilename(
            title="Load Progress",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    progress_data = json.load(f)
                
                
                for card in self.cards:
                    card_key = f"{card['order']}{card['name']}"
                    if card_key in progress_data['card_states']:
                        card['grabbed'] = progress_data['card_states'][card_key]
                        if card['grabbed']:
                            card['grabbed_var'].set(True)
                            self.grabbed_cards.add(card_key)
                        else:
                            card['grabbed_var'].set(False)
                            self.grabbed_cards.discard(card_key)
                
                
                self.display_cards()
                self.update_summary()
                
                messagebox.showinfo("Success", f"Progress loaded from {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Could not load progress: {e}")

def main():
    root = tk.Tk()
    app = PicklistGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
