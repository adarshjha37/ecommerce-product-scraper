import csv
import requests
from bs4 import BeautifulSoup


def scrape_products():
  url = "https://webscraper.io/test-sites/e-commerce/allinone/computers/laptops"
  headers = {
      "User-Agent": (
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
          " (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
      )
  }

  print(f"[INFO] Fetching product data from {url}...")
  response = requests.get(url, headers=headers)

  if response.status_code != 200:
    print(f"[ERROR] Could not connect. Status Code: {response.status_code}")
    return

  soup = BeautifulSoup(response.text, "html.parser")
  products = []

  items = soup.find_all("div", class_="thumbnail")

  for item in items:
    title = item.find("a", class_="title").get("title", "").strip()
    price = item.find("h4", class_="price").get_text().strip()
    description = item.find("p", class_="description").get_text().strip()

    products.append({
        "Product Name": title,
        "Price": price,
        "Description": description,
    })

  filename = "ecommerce_products.csv"
  with open(filename, mode="w", newline="", encoding="utf-8") as file:
    fieldnames = ["Product Name", "Price", "Description"]
    writer = csv.DictWriter(file, fieldnames=fieldnames)
    writer.writeheader()
    for prod in products:
      writer.writerow(prod)

  print(
      f"[SUCCESS] Scraped {len(products)} products! Saved to '{filename}'."
  )


if __name__ == "__main__":
  scrape_products()
  
