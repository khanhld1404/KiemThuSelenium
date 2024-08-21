import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def test_case_tc01(driver):
    result = {"Test Case": "Kiểm tra hiển thị khi nhấn vào một danh mục", "Status": "Fail", "Message": ""}
    
    driver.get("https://hoasenhome.vn")
    driver.maximize_window()
    
    menu = driver.find_element(By.CSS_SELECTOR, ".sec-menu-category")
    ActionChains(driver).move_to_element(menu).perform()
    time.sleep(2)
    
    try:
        category = driver.find_element(By.XPATH, "//a[contains(@href, '/chau-rua')]")  
        category.click()
        time.sleep(3)

        if driver.current_url != "https://hoasenhome.vn/":
            result["Status"] = "Pass"
            result["Message"] = "Đã chuyển sang một trang khác chứa các sản phẩm."
        else:
            result["Message"] = "Không có trang nào khác được mở."
    except Exception as e:
        result["Message"] = str(e)
    
    return result

def test_case_tc02(driver):
    result = {"Test Case": "Kiểm tra hiển thị khi nhấn vào một danh mục con", "Status": "Fail", "Message": ""}
    
    driver.get("https://hoasenhome.vn")
    driver.maximize_window()
    
    menu = driver.find_element(By.CSS_SELECTOR, ".sec-menu-category")
    ActionChains(driver).move_to_element(menu).perform()
    time.sleep(2)
    
    try:
        parent_category = driver.find_element(By.XPATH, "//a[contains(@href, '/ton-hoa-sen')]")  
        ActionChains(driver).move_to_element(parent_category).perform()
        time.sleep(2)
        sub_category = driver.find_element(By.XPATH, "//a[contains(@href, '/ton-hoa-sen/ton-lanh-mau')]")  
        sub_category.click()
        time.sleep(3)
        if driver.current_url != "https://hoasenhome.vn/":
            result["Status"] = "Pass"
            result["Message"] = "Đã chuyển sang một trang khác chứa các sản phẩm."
        else:
            result["Message"] = "Không có trang nào khác được mở."
    except Exception as e:
        result["Message"] = str(e)
    
    return result

def test_case_tc03(driver):
    result = {"Test Case": "Kiểm tra tốc độ tải của trang khi danh mục có nhiều loại sản phẩm", "Status": "Fail", "Message": ""}
    
    driver.get("https://hoasenhome.vn/")
    driver.maximize_window()
    
    menu = driver.find_element(By.CSS_SELECTOR, ".sec-menu-category")
    ActionChains(driver).move_to_element(menu).perform()
    time.sleep(2)
    
    try:
        category = driver.find_element(By.XPATH, "//a[contains(@href, '/ngoi')]")  
        category.click()
        time.sleep(2)
        start_time = time.time()
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "item.box-shadow")))
        
        end_time = time.time()
        load_time = end_time - start_time
        
        if load_time < 0.5:
            result["Status"] = "Pass"
            result["Message"] = f"Tốc độ tải hợp lý ({load_time} giây)."
        else:
            result["Message"] = f"Tốc độ tải quá chậm ({load_time} giây)."
    except Exception as e:
        result["Message"] = str(e)
    
    return result

def save_results_to_excel(results, file_name="KTDM_results.xlsx"):
    df = pd.DataFrame(results)
    df.to_excel(file_name, index=False, engine='openpyxl')

def main():
    driver = webdriver.Chrome()
    results = []

    try:
        results.append(test_case_tc01(driver))
        results.append(test_case_tc02(driver))
        results.append(test_case_tc03(driver))
    finally:
        driver.quit()
        save_results_to_excel(results)

if __name__ == "__main__":
    main()
