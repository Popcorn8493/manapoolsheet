# ManaPoolSheet

A Python tool for sellers to automate the creation of a fulfillment list from their ManaPool account.

This script connects to the ManaPool and Scryfall APIs to retrieve unshipped orders, collect detailed data for each item, and enrich it with card images. The primary goal is to create a consolidated CSV file that can be used for efficient picking and shipping of orders.

## Features

  - Fetches order data using the ManaPool API.
  - Automatically filters out already shipped orders to create a picklist of pending items.
  - Consolidates all items from all unshipped orders into a single CSV file.
  - Logs all operations and errors to a local file for troubleshooting.
  - Exports a detailed CSV with columns for order ID, quantity, name, set, condition, price, and image URL.

## Requirements

Python 3 

`pip install -r requirements.txt`

## Setup and Usage

**Provide Credentials**: In the same directory as the script, create two new files, using your api key from https://manapool.com/seller/integrations/manapool-api 

      - email.txt: Add your ManaPool account email to this file
      - api-key.txt: Add your ManaPool API token to this file

**Run the Script**: Open your terminal or command prompt, navigate to the script's directory, and run the following command:

    
    python manapoolsheet.py
   

When finished, you will find the `fulfillment_list.csv` and `order_retrieval_log.txt` files in the same directory.
