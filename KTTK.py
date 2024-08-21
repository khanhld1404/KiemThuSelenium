from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

# Khởi động WebDriver
driver = webdriver.Chrome()

def search_product(keyword, description):
    # B1: Mở trang web
    driver.get("https://hoasenhome.vn")
    
    # Lấy URL ban đầu
    initial_url = driver.current_url
    
    # B2: Kích vào thanh tìm kiếm
    search_box = driver.find_element(By.NAME, "searchValue")  # Sử dụng tên 'searchValue'
    
    # B3: Nhập từ khóa vào thanh tìm kiếm
    search_box.clear()
    search_box.send_keys(keyword)
    
    # B4: Nhấn vào biểu tượng tìm kiếm hoặc Enter
    search_box.send_keys(Keys.RETURN)
    
    # Chờ một khoảng thời gian để kết quả tải về
    time.sleep(5)

    # Lấy URL sau khi tìm kiếm
    current_url = driver.current_url
    
    # Kiểm tra URL để xem có thay đổi hay không
    if initial_url != current_url:
        # Kiểm tra số lượng các sản phẩm hiển thị
        products = driver.find_elements(By.CLASS_NAME, "item.box-shadow")
        num_products = len(products)
        if num_products > 0:
            result = {
                "Từ khóa": keyword,
                "Mô tả": description,
                "Thực hiện được tìm kiếm": "Có",
                "Tìm được kết quả": "Tìm được kết quả"
            }
        else:
            result = {
                "Từ khóa": keyword,
                "Mô tả": description,
                "Thực hiện được tìm kiếm": "Có",
                "Tìm được kết quả": "Không tìm được kết quả"
            }
    else:
        result = {
                "Từ khóa": keyword,
                "Mô tả": description,
                "Thực hiện được tìm kiếm": "Không",
                "Tìm được kết quả": "Không tìm được kết quả"
        }
    
    return result

def save_results_to_excel(results, file_name="KTTK_results.xlsx"):
        df = pd.DataFrame(results)
        df.to_excel(file_name, index=False, engine='openpyxl')

def main():
    # Danh sách từ khóa và mô tả nội dung kiểm thử
    search_keywords = [
        {"Từ khóa": "Bồn", "Mô tả": "Tìm kiếm sản phẩm với từ khóa trùng khớp một phần"},
        {"Từ khóa": "Con lợn", "Mô tả": "Tìm kiếm sản phẩm với tên không tồn tại"},
        {"Từ khóa": "vật liệu", "Mô tả": "Tìm kiếm sản phẩm bằng từ khóa chung chung"},
        {"Từ khóa": "Sơn14@", "Mô tả": "Tìm kiếm sản phẩm với từ khóa có chứa ký tự đặc biệt"},
        {"Từ khóa": "Quạt", "Mô tả": "Tìm kiếm sản phẩm bằng từ khóa có dấu"},
        {"Từ khóa": "Quat", "Mô tả": "Tìm kiếm sản phẩm bằng từ khóa không dấu"},
        {"Từ khóa": "a", "Mô tả": "Tìm kiếm sản phẩm bằng từ khóa chỉ có một kí tự"},
        {"Từ khóa": "abc", "Mô tả": "Tìm kiếm sản phẩm bằng từ khóa có nhiều kí tự ngẫu nhiên"},
        {"Từ khóa": "", "Mô tả": "Tìm kiếm sản phẩm bằng từ khóa trống"},
        {"Từ khóa": "Sơn", "Mô tả": "Tìm kiếm sản phẩm bằng từ khóa hợp lệ"}
    ]
    all_results = []
    for item in search_keywords:
        result = search_product(item["Từ khóa"], item["Mô tả"])
        all_results.append(result)
    driver.quit()
    save_results_to_excel(all_results)

if __name__ == "__main__":
    main()
