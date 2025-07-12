import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import time
import csv
import re

# Get today's date in the format YYYY-MM-DD
today_time_date = time.strftime("%Y-%m-%d")

# Function to safely get text from an element
def get_text_safe(element):
    try:
        return element.text.split("\n", 1)[-1].strip() if "\n" in element.text else element.text.strip()
    except AttributeError:
        return ""

# Function to extract data from a section
def extract_section_data(section):
    data = {}
    items = section.find_elements(By.CSS_SELECTOR, "li.list-item")
    for item in items:
        label = get_text_safe(item.find_element(By.CSS_SELECTOR, ".list-item-label"))
        value = get_text_safe(item.find_element(By.CSS_SELECTOR, ".list-item-content"))
        data[label] = value
    return data

# Function to process a single lot ID
def process_lot(lot, driver):
    try:
        # Open the website
        driver.get("https://montreal.ca/role-evaluation-fonciere")
        driver.maximize_window()

        # Locate and click the radio button
        div_element = driver.find_element(By.CSS_SELECTOR, "div[data-test='item-lot-renove']")
        div_element.click()

        # Scroll down in steps before clicking
        for _ in range(3):
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.PAGE_DOWN)
            time.sleep(0.5)

        # Click the submit button
        button = driver.find_element(By.CSS_SELECTOR, "button[data-test='submit']")
        button.click()
        time.sleep(0.5)

        # Locate and interact with the lot field
        lot_field = driver.find_element(By.CSS_SELECTOR, "input[class='form-control']")
        lot_field.clear()
        lot_field.send_keys(lot)
        time.sleep(1)

        # Click the search button
        button = driver.find_element(By.CSS_SELECTOR, "button[class='btn-squared btn btn-primary']")
        button.click()
        time.sleep(10)

        # Check if the "button[data-test='button']" element exists
        items = driver.find_elements(By.CSS_SELECTOR, "button[data-test='button']")
        if len(items) == 0:
            print(f"No details button found for lot {lot}. Saving to missed_value.csv.")
            with open("missed_value.csv", mode="a", newline="", encoding="utf-8") as missed_file:
                writer = csv.writer(missed_file)
                writer.writerow([lot])
            return None  # Skip this lot and move to the next one

        # Click the details button
        items[0].click()
        time.sleep(10)

        # Extract data from all sections
        identification_section = driver.find_element(By.CSS_SELECTOR, "section.layouts\\:stack h2#identification + ul")
        identification_data = extract_section_data(identification_section)

        proprietaires_section = driver.find_element(By.CSS_SELECTOR, "section.layouts\\:stack h2#proprietaires + ul")
        proprietaires_data = extract_section_data(proprietaires_section)

        terrain_section = driver.execute_script(terrain_section_script)
        terrain_data = extract_section_data(terrain_section)

        batiment_section = driver.execute_script(batiment_section_script)
        batiment_data = extract_section_data(batiment_section)

        role_courant_section = driver.execute_script(Rôle_courant_script)
        role_courant_data = extract_section_data(role_courant_section)

        role_anterieur_section = driver.execute_script(Rôle_antérieur_script)
        role_anterieur_data = extract_section_data(role_anterieur_section)

        repartition_section = driver.find_element(By.CSS_SELECTOR, "section.layouts\\:stack h2#repartition + ul")
        repartition_data = extract_section_data(repartition_section)
        
        Les_informations_script = """
        var elements = document.querySelectorAll('div.list-item-content');
        for (var i = 0; i < elements.length; i++) {
            var label = elements[i].querySelector('div.list-item-label');
            if (label && label.textContent.trim() === 'Les informations présentées dans ce rapport sont en date du') {
                return elements[i]; // Return the parent div containing the label and date
            }
        }
        return null;
        """

        # Execute the script and safely extract the text
        Les_informations = driver.execute_script(Les_informations_script)
        if Les_informations is not None:
            Les_informations_text = Les_informations.text
        else:
            Les_informations_text = "N/A"
            print(f"Warning: 'Les informations présentées dans ce rapport sont en date du' not found for lot {lot}.")

        # JavaScript script to find the element
        Date_du_script = """
        var elements = document.querySelectorAll('div.list-item-content');
        for (var i = 0; i < elements.length; i++) {
            var label = elements[i].querySelector('div.list-item-label');
            if (label && label.textContent.trim() === 'Date du rapport') {
                return elements[i]; // Return the parent div containing the label and date
            }
        }
        return null;
        """

        # Execute the script and safely extract the text
        Date_du = driver.execute_script(Date_du_script)
        if Date_du is not None:
            Date_du_text = Date_du.text
        else:
            Date_du_text = "N/A"
            print(f"Warning: 'Date du rapport' not found for lot {lot}.")
        # Combine all data into a single dictionary

        def extract_date(text):
            match = re.search(r'\d{4}-\d{2}-\d{2}', text)  # Matches YYYY-MM-DD format
            return match.group(0) if match else None  # Return only the date

        # Extract and clean the data
        Les_informations_date = extract_date(Les_informations.text)
        Date_du_date = extract_date(Date_du.text)

        all_data = {
            "ID": lot,
            "Identification": identification_data,
            "Propriétaire": proprietaires_data,
            "Caractéristiques du terrain": terrain_data,
            "Caractéristiques du bâtiment principal": batiment_data,
            "Rôle courant": role_courant_data,
            "Rôle antérieur": role_anterieur_data,
            "Répartition fiscale": repartition_data,
            "Les informations présentées dans ce rapport sont en date du": Les_informations_date,
            "Date du rapport": Date_du_date,
            "Date de l'extraction": today_time_date,
        }

        # Prepare the new row
        new_row = {
            "ID": all_data["ID"],
            **all_data["Identification"],
            **all_data["Propriétaire"],
            **all_data["Caractéristiques du terrain"],
            **all_data["Caractéristiques du bâtiment principal"],
            **all_data["Rôle courant"],
            **all_data["Rôle antérieur"],
            **all_data["Répartition fiscale"],
            "Les informations présentées dans ce rapport sont en date du": Les_informations_date,  # Store as string
            "Date du rapport": Date_du_date,  # Store as string
            "Date de l'extraction": all_data["Date de l'extraction"],
        }

        return new_row

    except Exception as e:
        print(f"Error processing lot {lot}: {e}")
        with open("error.csv", mode="a", newline="", encoding="utf-8") as error_file:
            writer = csv.writer(error_file)
            writer.writerow([lot])
        return None

    finally:
        # Navigate back to the main page instead of closing the browser
        driver.get("https://montreal.ca/role-evaluation-fonciere")

# JavaScript scripts for extracting sections
terrain_section_script = """
var h3Elements = document.querySelectorAll('h3.h4');
for (var i = 0; i < h3Elements.length; i++) {
    if (h3Elements[i].textContent.trim() === 'Caractéristiques du terrain') {
        return h3Elements[i].nextElementSibling;
    }
}
return null;
"""

batiment_section_script = """
var h3Elements = document.querySelectorAll('h3.h4');
for (var i = 0; i < h3Elements.length; i++) {
    if (h3Elements[i].textContent.trim() === 'Caractéristiques du bâtiment principal') {
        return h3Elements[i].nextElementSibling;
    }
}
return null;
"""

Rôle_courant_script = """
var h3Elements = document.querySelectorAll('h3.h4');
for (var i = 0; i < h3Elements.length; i++) {
    if (h3Elements[i].textContent.trim() === 'Rôle courant') {
        return h3Elements[i].nextElementSibling;
    }
}
return null;
"""

Rôle_antérieur_script = """
var h3Elements = document.querySelectorAll('h3.h4');
for (var i = 0; i < h3Elements.length; i++) {
    if (h3Elements[i].textContent.trim() === 'Rôle antérieur') {
        return h3Elements[i].nextElementSibling;
    }
}
return null;
"""

# Read lot IDs from the Excel file
excel_file_path = r"G:\New folder\testmontreal.xlsx"
df = pd.read_excel(excel_file_path)
lot_ids = df["NO_LOT"].tolist()

# Define the output Excel file
output_excel_file = r"G:\New folder\extracted_data.xlsx"
missed_value = r"G:\New folder\missed_value.csv"
error = r"G:\New folder\error.csv"

# Initialize missed_value.csv and error.csv if they don't exist
if not os.path.isfile(missed_value):
    with open(missed_value, mode="a", newline="", encoding="utf-8") as missed_file:
        writer = csv.writer(missed_file)
        writer.writerow(["NO_LOT"])

if not os.path.isfile(error):
    with open(error, mode="a", newline="", encoding="utf-8") as error_file:
        writer = csv.writer(error_file)
        writer.writerow(["NO_LOT"])

# Set up Chrome driver
options = webdriver.ChromeOptions()
options.add_argument("--incognito")
chrome_driver_path = ChromeDriverManager().install()
service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=options)

# Process each lot ID
all_rows = []
for lot in lot_ids:
    try:
        new_row = process_lot(lot, driver)
        
        # If no details button was found or an error occurred, skip to the next lot
        if new_row is None:
            continue
        
        # Append the new row to the list
        all_rows.append(new_row)
        print(f"Data for lot {lot} processed.")
    
    except Exception as e:
        print(f"Unexpected error processing lot {lot}: {e}")
        with open(error, mode="a", newline="", encoding="utf-8") as error_file:
            writer = csv.writer(error_file)
            writer.writerow([lot])

# Save all rows to the Excel file
if all_rows:
    # Read existing data from the Excel file (if it exists)
    if os.path.isfile(output_excel_file):
        existing_df = pd.read_excel(output_excel_file)
    else:
        existing_df = pd.DataFrame()

    # Append new rows to the existing data
    updated_df = pd.concat([existing_df, pd.DataFrame(all_rows)], ignore_index=True)

    # Save the updated DataFrame to the Excel file
    updated_df.to_excel(output_excel_file, index=False)
    print(f"All data saved to {output_excel_file}")
else:
    print("No data to save.")

# Close the browser after processing all lots
driver.quit()