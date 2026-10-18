import re
import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

print("[INFO] Fetching complete product data from e-commerce site...")
url = "https://webscraper.io/test-sites/e-commerce/allinone/computers/laptops"
response = requests.get(url)
soup = BeautifulSoup(response.text, 'html.parser')

products = []
for item in soup.select('.thumbnail'):
    # Product Title & Link
    title_elem = item.select_one('.title')
    title = title_elem.get('title', '').strip() if title_elem and title_elem.get('title') else (title_elem.text.strip() if title_elem else 'N/A')
    
    # Price
    price_elem = item.select_one('.price')
    price_str = price_elem.text.strip() if price_elem else '$0.00'
    try:
        price_num = float(price_str.replace('$', '').replace(',', ''))
    except ValueError:
        price_num = 0.0

    # Description
    desc_elem = item.select_one('.description')
    description = desc_elem.text.strip() if desc_elem else ''

    # Ratings & Reviews
    reviews_elem = item.select_one('.ratings .pull-right')
    reviews_text = reviews_elem.text.strip() if reviews_elem else '0 reviews'
    reviews_count = int(re.search(r'\d+', reviews_text).group()) if re.search(r'\d+', reviews_text) else 0
    
    stars = len(item.select('.ratings .glyphicon-star'))

    # Extract Specs from Description
    screen_match = re.search(r'(\d+(\.\d+)?[\"\'](\s*IPS)?)', description, re.IGNORECASE)
    screen_size = screen_match.group(0) if screen_match else "N/A"

    ram_match = re.search(r'(\d+\s*GB)', description, re.IGNORECASE)
    ram = ram_match.group(0) if ram_match else "N/A"

    storage_match = re.search(r'(\d+\s*(GB|TB)\s*(SSD|HDD)?)', description, re.IGNORECASE)
    storage = storage_match.group(0) if storage_match else "N/A"

    products.append({
        "Category": "Laptops",
        "Product Name": title,
        "Price ($)": price_num,
        "Rating": f"★ {stars}.0 / 5",
        "Total Reviews": reviews_count,
        "Screen Size": screen_size,
        "RAM": ram,
        "Storage": storage,
        "Full Description": description
    })

df = pd.DataFrame(products)
excel_file = "ecommerce_products_perfect.xlsx"

# Export to Excel with Professional Styling
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Laptop Catalog')
    workbook = writer.book
    worksheet = writer.sheets['Laptop Catalog']

    # Gridlines
    worksheet.views.sheetView[0].showGridLines = True

    # Styles Definition
    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")  # Deep Navy Blue
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    
    alt_row_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")  # Soft Light Blue Tint
    white_row_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    data_font = Font(name="Segoe UI", size=10)
    bold_data_font = Font(name="Segoe UI", size=10, bold=True)
    
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )

    # Header Row
    worksheet.row_dimensions[1].height = 28
    for col_idx in range(1, len(df.columns) + 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    # Data Rows
    for row_idx in range(2, len(df) + 2):
        worksheet.row_dimensions[row_idx].height = 22
        current_fill = alt_row_fill if row_idx % 2 == 0 else white_row_fill

        for col_idx in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.fill = current_fill
            cell.border = thin_border

            # Alignment & Number Formatting
            col_name = df.columns[col_idx - 1]
            if col_name == "Price ($)":
                cell.number_format = '"$"#,##0.00'
                cell.alignment = right_align
                cell.font = bold_data_font
            elif col_name in ["Category", "Rating", "Total Reviews", "Screen Size", "RAM", "Storage"]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    # Auto Column Width Adjustment
    for col in worksheet.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val = str(cell.value or '')
            if cell.number_format and '$' in cell.number_format:
                val = f"${cell.value:,.2f}" if isinstance(cell.value, (int, float)) else val
            if len(val) > max_len:
                max_len = len(val)
        
        # Clamp widths between 12 and 55
        adjusted_width = min(max(max_len + 3, 12), 55)
        worksheet.column_dimensions[col_letter].width = adjusted_width

print(f"[SUCCESS] Scraped {len(products)} products with complete details!")
print(f"[SUCCESS] Perfect Excel generated: '{excel_file}'")
