import os
import time
import glob  # 导入glob模块用于文件模式匹配
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 定义下载目录
# Define the download directory
parent_dir = os.path.dirname(os.getcwd())

# 在上一级目录下创建 downloads 文件夹路径
download_dir = os.path.join(parent_dir, "downloads")

# 如果路径不存在就创建
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

print(f"文件将下载到: {download_dir}")

# ========== 1. 浏览器配置 ==========
# Browser configuration
chrome_options = Options()
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--ignore-ssl-errors')
chrome_options.add_argument('--allow-insecure-localhost')

# 配置下载偏好设置，让Chrome自动下载文件到指定目录
# Configure download preferences for Chrome to automatically download files to the specified directory
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,  # 不弹出下载确认框
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True  # 禁用安全浏览，以防下载被阻止
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 30)

# ========== 2. 打开登录页面 ==========
# Open the login page
# 修正后的登录URL
# Corrected login URL
driver.get('https://172.31.254.244:5443/#/user/login')

# ========== 3. 输入用户名和密码 ==========
# Enter username and password
username_input = wait.until(EC.presence_of_element_located((By.ID, 'username')))
# 修正后的用户名
# Corrected username
username_input.send_keys('123')

password_input = driver.find_element(By.ID, 'password')
# 修正后的密码
# Corrected password
password_input.send_keys('a@123456789')

# 点击登录按钮
# Click the login button
login_button = driver.find_element(By.XPATH, '//button[span[contains(text(), "登 录")]]')
login_button.click()

# ========== 4. 等待登录成功页面 ==========
# Wait for the successful login page
# 假设登录成功后URL包含 '/#/threat-awareness/overview'
# Assuming the URL after successful login contains '/#/threat-awareness/overview'
wait.until(EC.url_contains('/#/threat-awareness/overview'))
print("登录成功，已跳转到威胁感知页面。")

# ========== 5. 直接访问目标页面 ==========
# Directly access the target page
driver.get('https://172.31.254.244:5443/#/threat-source/event-trace')
print("已跳转到事件溯源页面。")

# ========== 6. 切换到 iframe ==========
# Switch to the iframe
# 等待iframe出现并切换
# Wait for the iframe to be present and switch to it
iframes = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'iframe')))
driver.switch_to.frame(iframes[0])
print("已切换到iframe。")

# ========== 7. 等待 canvas 遮挡物消失 ==========
# Wait for the canvas overlay to disappear
# 等待canvas遮挡物消失
# Wait for the canvas overlay to disappear
wait.until(EC.invisibility_of_element_located((By.TAG_NAME, 'canvas')))
print("Canvas遮挡物已消失。")

# ========== 8. 点击“全部导出”按钮 ==========
# Click the "全部导出" (Export All) button
export_button = wait.until(
    EC.element_to_be_clickable((By.XPATH, '//button[span[contains(text(), "全部导出")]]'))
)
driver.execute_script("arguments[0].scrollIntoView();", export_button)
driver.execute_script("arguments[0].click();", export_button)
print("已点击“全部导出”按钮。")

# ========== 9. 点击 XLSX 菜单 ==========
# Click the XLSX menu item
# 修改为点击XLSX按钮
# Changed to click the XLSX button
xlsx_button = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, '//li[contains(@class, "el-dropdown-menu__item") and contains(text(), "XLSX")]'))
)
driver.execute_script("arguments[0].click();", xlsx_button)
print("已点击“XLSX”菜单项，开始下载。")

# ========== 10. 等待文件下载完成 ==========
# Wait for the file download to complete
# 等待文件出现在下载目录中，并确保其完成下载
# Wait for the file to appear in the download directory and ensure it completes downloading
timeout = 60  # 秒
start_time = time.time()
downloaded_file_name = None

print(f"等待文件下载到: {download_dir}")

while time.time() - start_time < timeout:
    # 查找所有 .crdownload 文件
    # Find all .crdownload files
    crdownload_files = glob.glob(os.path.join(download_dir, '*.crdownload'))

    # 查找所有 .xlsx 文件
    # Find all .xlsx files
    xlsx_files = glob.glob(os.path.join(download_dir, '*.xlsx'))

    if xlsx_files:
        # 如果有 .xlsx 文件，假设是下载完成的文件
        # If there are .xlsx files, assume one is the completed download
        downloaded_file_name = os.path.basename(xlsx_files[0])  # 取第一个找到的xlsx文件
        print(f"检测到已完成的XLSX文件: {downloaded_file_name}")
        break
    elif crdownload_files:
        # 如果有 .crdownload 文件，等待其完成
        # If there are .crdownload files, wait for them to complete
        temp_file_path = crdownload_files[0]  # 取第一个找到的crdownload文件
        print(f"检测到临时下载文件: {os.path.basename(temp_file_path)}，等待完成...")
        prev_size = -1
        # 等待文件大小不再变化，或者文件消失（被重命名）
        # Wait for file size to stop changing, or for the file to disappear (be renamed)
        while time.time() - start_time < timeout:
            try:
                if not os.path.exists(temp_file_path):
                    # 如果临时文件消失了，说明可能已经重命名，再次查找xlsx文件
                    # If the temporary file disappears, it might have been renamed, search for xlsx file again
                    xlsx_files_after_rename = glob.glob(os.path.join(download_dir, '*.xlsx'))
                    if xlsx_files_after_rename:
                        downloaded_file_name = os.path.basename(xlsx_files_after_rename[0])
                        print(f"临时文件消失，找到最终XLSX文件: {downloaded_file_name}")
                        break

                current_size = os.path.getsize(temp_file_path)
                if current_size == prev_size and current_size > 0:
                    print(f"临时文件 {os.path.basename(temp_file_path)} 大小稳定，但仍是.crdownload，继续等待重命名...")
                    # 文件大小稳定但仍是.crdownload，可能是下载完成但未重命名，或需要更多时间
                    # File size is stable but still .crdownload, might be downloaded but not renamed, or needs more time
                    time.sleep(2)  # 稍微多等一下
                prev_size = current_size
                time.sleep(1)
            except FileNotFoundError:
                # 文件未找到，可能已经完成并重命名，再次检查xlsx文件
                # File not found, might have completed and been renamed, check for xlsx file again
                xlsx_files_after_rename = glob.glob(os.path.join(download_dir, '*.xlsx'))
                if xlsx_files_after_rename:
                    downloaded_file_name = os.path.basename(xlsx_files_after_rename[0])
                    print(f"临时文件消失，找到最终XLSX文件: {downloaded_file_name}")
                    break
                else:
                    print(f"文件 {os.path.basename(temp_file_path)} 暂时未找到，继续等待...")
                time.sleep(1)
            except Exception as e:
                print(f"检查文件大小时发生错误: {e}")
                time.sleep(1)
        if downloaded_file_name:
            break  # 如果找到了最终文件，跳出外层循环

    time.sleep(1)  # 每秒检查一次

if downloaded_file_name:
    print(f"文件 '{downloaded_file_name}' 已下载到 '{download_dir}'。")
else:
    print("文件下载超时或未检测到最终XLSX文件。")

driver.quit()
