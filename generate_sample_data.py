# generate_sample_data.py
# ─────────────────────────────────────────────────────────────
# Run this file ONCE to create a realistic sample Excel dataset.
# This simulates a Bangladesh-based online shop's 6-month sales.
# After running, you'll find: data/sample_sales.xlsx
# ─────────────────────────────────────────────────────────────

import pandas as pd
import numpy as np
import random
import os
from datetime import datetime, timedelta

# ── Set a seed so results are reproducible ──────────────────
random.seed(42)
np.random.seed(42)

# ── 1. Define your business data ────────────────────────────

PRODUCTS = [
    {"name": "Cotton Panjabi (White)",    "price": 850},
    {"name": "Cotton Panjabi (Blue)",     "price": 950},
    {"name": "Silk Saree (Red)",          "price": 2200},
    {"name": "Silk Saree (Green)",        "price": 2500},
    {"name": "Kids Dress (3-5 yrs)",      "price": 480},
    {"name": "Kids Dress (6-9 yrs)",      "price": 550},
    {"name": "Formal Shirt (Men)",        "price": 700},
    {"name": "Casual T-Shirt",            "price": 350},
    {"name": "Ladies Kurti (Cotton)",     "price": 620},
    {"name": "Ladies Kurti (Printed)",    "price": 750},
    {"name": "Winter Jacket (Men)",       "price": 1800},
    {"name": "Shawl / Orna",             "price": 400},
]

CUSTOMERS = [
    {"name": "Rahim Uddin",      "phone": "01711-111001"},
    {"name": "Fatema Begum",     "phone": "01812-222002"},
    {"name": "Karim Hossain",    "phone": "01911-333003"},
    {"name": "Sumaiya Akter",    "phone": "01615-444004"},
    {"name": "Jamal Sheikh",     "phone": "01722-555005"},
    {"name": "Nasrin Jahan",     "phone": "01833-666006"},
    {"name": "Tofazzal Islam",   "phone": "01944-777007"},
    {"name": "Roksana Khanam",   "phone": "01555-888008"},
    {"name": "Belal Ahmed",      "phone": "01766-999009"},
    {"name": "Shirin Sultana",   "phone": "01877-000010"},
    {"name": "Monir Hossain",    "phone": "01988-111011"},
    {"name": "Parvin Akter",     "phone": "01600-222012"},
    {"name": "Sohel Rana",       "phone": "01711-333013"},
    {"name": "Lovely Begum",     "phone": "01812-444014"},
    {"name": "Rubel Mia",        "phone": "01933-555015"},
    {"name": "Champa Khatun",    "phone": "01644-666016"},
    {"name": "Liton Das",        "phone": "01755-777017"},
    {"name": "Mitu Akter",       "phone": "01866-888018"},
    {"name": "Khairul Islam",    "phone": "01977-999019"},
    {"name": "Sharmin Nahar",    "phone": "01588-000020"},
    # Some one-time / rare customers
    {"name": "Abdul Kader",      "phone": "01700-123456"},
    {"name": "Renu Begum",       "phone": "01800-234567"},
    {"name": "Sajib Hasan",      "phone": "01900-345678"},
    {"name": "Rina Akter",       "phone": "01600-456789"},
    {"name": "Noman Ali",        "phone": "01700-567890"},
]

PAYMENT_METHODS = ["bKash", "Nagad", "Cash on Delivery", "Rocket", "Bank Transfer"]
DELIVERY_STATUSES = ["Delivered", "Delivered", "Delivered", "Returned", "Pending"]
# Note: "Delivered" appears 3x so it's more likely (realistic)

# ── 2. Generate orders over 6 months ────────────────────────

start_date = datetime(2024, 7, 1)   # Start: July 2024
end_date   = datetime(2024, 12, 31) # End:   December 2024
date_range = (end_date - start_date).days

orders = []
order_id_counter = 10001

for _ in range(500):  # Generate 500 orders

    # Pick a random date (slightly more orders on weekends — realistic!)
    random_day   = random.randint(0, date_range)
    order_date   = start_date + timedelta(days=random_day)

    # Pick a customer (top 15 customers appear more often — loyalty!)
    if random.random() < 0.75:
        customer = random.choice(CUSTOMERS[:15])   # loyal customers
    else:
        customer = random.choice(CUSTOMERS[15:])   # one-time customers

    # Pick a product
    product = random.choice(PRODUCTS)

    # Quantity (mostly 1-2 items, rarely 3-4)
    quantity = random.choices([1, 2, 3, 4], weights=[60, 25, 10, 5])[0]

    # Small price variation (discounts, negotiations)
    price_variation = random.uniform(0.90, 1.05)
    unit_price      = round(product["price"] * price_variation, 0)
    total_revenue   = unit_price * quantity

    # Payment & delivery
    payment_method   = random.choice(PAYMENT_METHODS)
    delivery_status  = random.choice(DELIVERY_STATUSES)

    orders.append({
        "Order ID":        f"ORD-{order_id_counter}",
        "Order Date":      order_date.strftime("%Y-%m-%d"),
        "Customer Name":   customer["name"],
        "Customer Phone":  customer["phone"],
        "Product Name":    product["name"],
        "Quantity":        quantity,
        "Unit Price (BDT)": unit_price,
        "Total Revenue (BDT)": total_revenue,
        "Payment Method":  payment_method,
        "Delivery Status": delivery_status,
    })

    order_id_counter += 1

# ── 3. Add some intentional "messy" data ────────────────────
# (So our dashboard can handle real-world imperfect data)

# A few missing customer names
for i in random.sample(range(500), 10):
    orders[i]["Customer Name"] = None

# A few blank delivery statuses
for i in random.sample(range(500), 5):
    orders[i]["Delivery Status"] = None

# ── 4. Create DataFrame and save to Excel ───────────────────

df = pd.DataFrame(orders)

# Sort by date (most recent first — looks cleaner)
df["Order Date"] = pd.to_datetime(df["Order Date"])
df = df.sort_values("Order Date").reset_index(drop=True)

# Make sure the data/ folder exists
os.makedirs("data", exist_ok=True)

# Save to Excel
output_path = "data/sample_sales.xlsx"
df.to_excel(output_path, index=False, engine="openpyxl")

print("=" * 55)
print("✅ Sample data generated successfully!")
print(f"📁 File saved at: {output_path}")
print(f"📊 Total orders:  {len(df)}")
print(f"👥 Unique customers: {df['Customer Name'].nunique()}")
print(f"📦 Unique products:  {df['Product Name'].nunique()}")
print(f"💰 Total revenue: {df['Total Revenue (BDT)'].sum():,.0f} BDT")
print("=" * 55)
print("\n📋 First 5 rows preview:")
print(df.head())