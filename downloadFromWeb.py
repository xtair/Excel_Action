import time
import pickle
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException


def openAndDownloads():
    driver = webdriver.Chrome()
    driver.get('https://accounts.zoho.com/signin')  # replace with your login page URL

    # Locate and click the login
    login_button = driver.find_element('id', 'nextbtn')

    # Find the username and password fields. This assumes they're identified by their name attributes.
    # Input credentials into the fields
    wait = WebDriverWait(driver, 10)
    username_field = wait.until(EC.element_to_be_clickable((By.ID, 'login_id')))
    username_field.send_keys('xuanzhen.tai@happy-distro.com')

    # submit button
    login_button.click()

    wait = WebDriverWait(driver, 10)
    password_field = wait.until(EC.element_to_be_clickable((By.ID, 'password')))
    if password_field.is_displayed():
        password_field.send_keys('%fhn32xZvjS2AQc')
    else:
        print("Element is not visible")

    login_button.click()
    print("login success")

    # after login
    # time.sleep(10)
    driver.get('https://books.zoho.com/Home.jsp#/organizations')

    #wait.until(EC.presence_of_element_located((By.CLASS_NAME, "zProName")))


    while True:
        time.sleep(10)

    # options = webdriver.ChromeOptions()
    # options.add_experimental_option("detach", True)
    # driver = webdriver.Chrome(options=options)

    """
    # 创建一个新的浏览器实例
    driver = webdriver.Chrome()

    # 打开一个网页
    driver.get('https://books.zoho.com/app/764770972#/reports/inventorysummary?'
               'item_ids=2943539000030240008&report_date=2023-10-26&show_actual_stock=true')

    # 找到并点击一个按钮
    button = driver.find_element_by_xpath('//button[@class="dropdown-item pl-4"][contains(text(), "XLSX (Microsoft Excel)")]')
    button.click()

    # 停留几秒，确保文件下载完毕（这只是一个简单的例子，实际上你可能需要其他的方法确保文件已下载）
    import time
    time.sleep(5)

    # 关闭浏览器
    driver.quit()
    """
