import cv2
import pytesseract
import pandas as pd
import time
import os
import re
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Folder to watch for new invoices
FOLDER_TO_WATCH = "invoices"
# Output Excel file
EXCEL_FILE = "invoices.xlsx"

# Ensure the folder exists
if not os.path.exists(FOLDER_TO_WATCH):
    os.makedirs(FOLDER_TO_WATCH)

# Function to extract text from an invoice image
def extract_text_from_image(image_path):
    img = cv2.imread(image_path)
    text = pytesseract.image_to_string(img, lang="eng+rus+heb")  # Supports English, Russian, and Hebrew
    return text

# Function to extract invoice description and total sum
def extract_invoice_details(text):
    description = "Unknown Invoice"
    
    # Extract invoice description (first line)
    lines = text.split("\n")
    if lines:
        description = lines[0].strip()

    # Extract total sum (assuming pattern like "Total: 123.45")
    match = re.search(r"Total[:\s]+(\d+\.\d{2})", text, re.IGNORECASE)
    total_sum = float(match.group(1)) if match else None

    return description, total_sum

# Function to update Excel with invoice data
def update_excel(file_path, image_name, description, total_sum):
    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Image Name", "Description", "Total Sum"])
    
    new_data = pd.DataFrame([[image_name, description, total_sum]], columns=df.columns)
    df = pd.concat([df, new_data], ignore_index=True)
    
    df.to_excel(file_path, index=False)
    print(f"Updated Excel with: {image_name} - {description} - {total_sum}")

# Watchdog event handler to monitor folder changes
class InvoiceHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        if event.src_path.lower().endswith(('.png', '.jpg', '.jpeg')):
            time.sleep(2)  # Wait for the file to be fully written
            print(f"New invoice detected: {event.src_path}")
            text = extract_text_from_image(event.src_path)
            description, total_sum = extract_invoice_details(text)
            update_excel(EXCEL_FILE, os.path.basename(event.src_path), description, total_sum)

# Start monitoring folder
observer = Observer()
event_handler = InvoiceHandler()
observer.schedule(event_handler, FOLDER_TO_WATCH, recursive=False)
observer.start()

print("Invoice watcher is running... Press Ctrl+C to stop.")

try:
    while True:
        time.sleep(10)
except KeyboardInterrupt:
    print("\nStopping invoice watcher...")
    observer.stop()

observer.join()
