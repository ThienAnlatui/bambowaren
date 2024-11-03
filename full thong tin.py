import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from bs4 import BeautifulSoup
import time
import pandas as pd

# Cookie setup
cookie_string = """uaid=AYqPI9-e39fpZGTGAIG97h1DJKZjZACCdNl5FTC6Wqk0MTNFyUrJzyPdNzfD2MfRN8fXwrAiPDs8JDLINSQ-KdkoXKmWAQA.;"""
cookies = [{"name": item.split('=')[0].strip(), "value": item.split('=')[1].strip(), "domain": "www.etsy.com"} for item in cookie_string.split(";") if "=" in item]

def scrape_etsy_from_url(url, items):
    options = uc.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = uc.Chrome(options=options)
    
    try:
        driver.get(url)
        time.sleep(3)
        
        for cookie in cookies:
            driver.add_cookie(cookie)
        
        driver.refresh()
        time.sleep(5)
        
        current_page = 1
        while True:
            soup = BeautifulSoup(driver.page_source, "html.parser")
            products = soup.select(".v2-listing-card")
            
            print(f"Trang {current_page}: {len(products)} sản phẩm được tìm thấy.")
            
            if not products:
                print("Không có sản phẩm nào được tìm thấy trên trang này.")
                break
            
            for product in products:
                try:
                    # Lấy các thông tin chi tiết sản phẩm
                    product_id = product.get("data-listing-id", "N/A")
                    title = product.select_one(".v2-listing-card__title").text.strip() if product.select_one(".v2-listing-card__title") else "N/A"
                    image_url = product.select_one("img")["src"] if product.select_one("img") else "N/A"
                    product_url =  product.select_one("a").get("href", "") if product.select_one("a") else "N/A"
                    
                    seller = product.select_one(".v2-listing-card__shop a")
                    seller_name = seller.text.strip() if seller else "N/A"
                    seller_url =  seller.get("href", "") if seller else "N/A"
                    
                    # Giá gốc và giá giảm
                    price_original = product.select_one(".currency-value.text-strike")
                    price_discounted = product.select_one(".currency-value:not(.text-strike)")
                    
                    price_original = price_original.text.strip() if price_original else "N/A"
                    price_discounted = price_discounted.text.strip() if price_discounted else "N/A"
                    
                    # Tính phần trăm giảm giá
                    discount_percent = "N/A"
                    if price_original != "N/A" and price_discounted != "N/A":
                        try:
                            price_original_value = float(price_original.replace(",", ""))
                            price_discounted_value = float(price_discounted.replace(",", ""))
                            discount_percent = f"{round((1 - price_discounted_value / price_original_value) * 100, 2)}%"
                        except ValueError:
                            discount_percent = "N/A"
                    
                    # Đánh giá và lượt đánh giá
                    rating = product.select_one("span.screen-reader-only")
                    rating = rating.text.strip() if rating else "N/A"
                    review_count = product.select_one(".wt-text-caption")
                    review_count = review_count.text.strip() if review_count else "N/A"
                    
                    items.append({
                        "ID": product_id,
                        "Tên sản phẩm": title,
                        "Ảnh URL": image_url,
                        "URL sản phẩm": product_url,
                        "Người bán": seller_name,
                        "URL người bán": seller_url,
                        "Giá gốc": price_original,
                        "Giá giảm": price_discounted,
                        "Phần trăm giảm": discount_percent,
                        "Đánh giá": rating,
                        "Lượt đánh giá": review_count
                    })
                
                except Exception as e:
                    print("Lỗi khi thu thập thông tin:", e)
            
            print(f"Số lượng sản phẩm sau khi thu thập: {len(items)}")
            
            try:
                time.sleep(3)
                next_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "#content > div > div.wt-bg-white.wt-grid__item-md-12.wt-pl-xs-1.wt-pr-xs-0.wt-pr-md-1.wt-pl-lg-0.wt-pr-lg-0.wt-bb-xs-1 > div > div.wt-mt-xs-3.wt-text-black > div.wt-grid.wt-pl-xs-0.wt-pr-xs-0.search-listings-group > div:nth-child(2) > div.wt-mb-xs-5.wt-mt-xs-6 > div > div > div > div.wt-hide-xs.wt-show-lg > nav > div > div:last-child > a"))
                )
                
                if "disabled" in next_button.get_attribute("class"):
                    print("Đã đến trang cuối.")
                    break
                
                next_button.click()
                current_page += 1
                time.sleep(15)
            
            except (NoSuchElementException, TimeoutException):
                print("Không tìm thấy nút tiếp theo hoặc thời gian chờ hết hạn.")
                break

    finally:
        driver.quit()  # Đảm bảo trình duyệt luôn đóng khi kết thúc
        print("Đang lưu dữ liệu vào tệp Excel...")
        if items:
            df = pd.DataFrame(items)
            df.to_excel("C:/Users/ACER/Desktop/jj3.xlsx", index=False)
            print("Đã lưu thông tin sản phẩm vào jj3.xlsx")
        else:
            print("Không có sản phẩm nào để lưu.")

def main():
    items = []
    url = "https://www.etsy.com/search?q=woven%20bamboo&ref=search_bar"
    print("Đang thu thập dữ liệu từ:", url)
    scrape_etsy_from_url(url, items)

if __name__ == "__main__":
    main()
