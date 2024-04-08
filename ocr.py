import pandas as pd
from selenium import webdriver
from PIL import Image
import pytesseract
import io
from collections import defaultdict
from time import sleep

# Set the path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Function to parse the OCR results into a list of dictionaries
def parse_text(ocr_text):
    parsed_results = []
    lines = ocr_text.split('\n')
    key = None
    for line in lines:
        if not key:
            key = line.strip()
        else:
            parsed_results.append((key, line.strip()))
            key = None  # Reset the key for the next pair
    return parsed_results

# Function to take a screenshot, crop to the desired area, and return the OCR results
def get_ocr_results(url, driver, left, top, right, bottom):
    driver.get(url)
    sleep(5)  # Wait for the page to load
    screenshot = driver.get_screenshot_as_png()  # Take screenshot
    image = Image.open(io.BytesIO(screenshot))
    
    # Crop to the specific area
    cropped_image = image.crop((left, top, right, bottom))
    
    # Perform OCR on the cropped area
    ocr_text = pytesseract.image_to_string(cropped_image)
    return ocr_text

# Setup Chrome WebDriver
driver = webdriver.Chrome()

# Read the Excel sheet with URLs
input_file = 'C:/Users/newpu/Downloads/Book.xlsx'
df = pd.read_excel(input_file)
urls = df['models'].tolist()  # Assuming 'models' column has URLs

# Dictionary to store parsed data for each URL
parsed_data_dict = defaultdict(dict)

# Coordinates for the specific area to crop before OCR
left = 300
top = 100
right = 1200
bottom = 700

# Iterate over the URLs, take screenshots, crop to the specific area, and parse the text
for url in urls:
    try:
        ocr_text = get_ocr_results(url, driver, left, top, right, bottom)
        key_value_pairs = parse_text(ocr_text)
        for key, value in key_value_pairs:
            parsed_data_dict[url][key] = value
    except Exception as e:
        print(f"Error processing URL: {url}. Error: {str(e)}")

# Cleanup and close the browser
driver.quit()

# Convert the parsed_data_dict to a DataFrame
output_df = pd.DataFrame.from_dict(parsed_data_dict, orient='index')

# Ensure the DataFrame is sorted by column headers to maintain consistency
output_df = output_df.reindex(sorted(output_df.columns), axis=1)

# Save the output DataFrame to a new Excel file
output_file = 'C:/Users/newpu/Downloads/output.xlsx'
output_df.to_excel(output_file, index_label='URL')

print("Data saved to Excel successfully!")
