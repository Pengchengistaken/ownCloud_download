"""
SharePoint 自动化下载工具
使用 Selenium ChromeDriver 自动化访问 SharePoint 共享文件夹并下载文件
"""

import os
import time
import logging
from pathlib import Path
from typing import List, Tuple, Optional
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
import config

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('sharepoint_download.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 全局统计
download_stats = {
    'total_files': 0,
    'existing_files': 0,
    'downloaded': 0,
    'failed': 0
}

class SharePointDownloader:
    def __init__(self):
        self.driver = self.setup_chrome_driver()
        self.wait = WebDriverWait(self.driver, config.SHAREPOINT_PAGE_LOAD_TIMEOUT)

    def setup_chrome_driver(self) -> webdriver.Chrome:
        """初始化 ChromeDriver"""
        chrome_options = Options()
        download_path = os.path.abspath(config.SHAREPOINT_DOWNLOAD_DIR)
        os.makedirs(download_path, exist_ok=True)
        
        prefs = {
            "download.default_directory": download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "safebrowsing.disable_download_protection": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_page_load_timeout(config.SHAREPOINT_PAGE_LOAD_TIMEOUT)
        
        logger.info(f"SharePoint ChromeDriver已初始化，下载路径: {download_path}")
        return driver

    def wait_for_page_load(self, timeout: int = 20):
        """等待页面加载完成"""
        try:
            self.wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            time.sleep(2)
        except TimeoutException:
            logger.warning("页面加载超时")

    def access_share_link(self, url: str) -> bool:
        """访问分享链接"""
        try:
            logger.info(f"正在访问 SharePoint 链接: {url}")
            self.driver.get(url)
            self.wait_for_page_load()
            
            # 检查是否进入了文件列表页面
            # SharePoint 匿名分享通常直接进入，如果有密码会停留在输入框
            if "guestaccess.aspx" in self.driver.current_url or "onedrive.aspx" in self.driver.current_url or "sharepoint.com" in self.driver.current_url:
                logger.info("成功到达目标页面")
                return True
            return False
        except Exception as e:
            logger.error(f"访问失败: {str(e)}")
            return False

    def scroll_to_load_all_files(self):
        """滚动加载所有文件（SharePoint 使用虚拟滚动）"""
        try:
            # 查找滚动容器，通常是带有 role='grid' 的 div
            container_selectors = ["[role='grid']", ".ms-List-page", ".od-ItemsList"]
            container = None
            for selector in container_selectors:
                elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                if elements:
                    container = elements[0]
                    break
            
            if not container:
                # 如果找不到容器，直接滚动 window
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                return

            last_height = self.driver.execute_script("return arguments[0].scrollHeight", container)
            while True:
                self.driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", container)
                time.sleep(2)
                new_height = self.driver.execute_script("return arguments[0].scrollHeight", container)
                if new_height == last_height:
                    break
                last_height = new_height
            
            # 滚回顶部
            self.driver.execute_script("arguments[0].scrollTop = 0", container)
            time.sleep(1)
        except Exception as e:
            logger.warning(f"滚动加载失败: {str(e)}")

    def get_items(self) -> List[dict]:
        """获取当前页面的所有项目（文件或文件夹）"""
        items = []
        try:
            # SharePoint 常见的行选择器
            row_selector = "div[role='row'][data-automationid='DetailsRow']"
            rows = self.driver.find_elements(By.CSS_SELECTOR, row_selector)
            
            for row in rows:
                try:
                    # 获取名称
                    name_btn = row.find_element(By.CSS_SELECTOR, "button[data-automationid='FieldRenderer-name']")
                    name = name_btn.text.strip()
                    
                    # 判断类型
                    is_folder = False
                    try:
                        # 检查是否有文件夹图标
                        row.find_element(By.CSS_SELECTOR, "i[data-icon-name='FabricFolder'], i[data-icon-name='FolderInverse']")
                        is_folder = True
                    except:
                        pass
                    
                    items.append({
                        'name': name,
                        'is_folder': is_folder,
                        'element': row
                    })
                except Exception as row_error:
                    logger.debug(f"解析行出错: {str(row_error)}")
                    continue
        except Exception as e:
            logger.error(f"获取列表失败: {str(e)}")
        
        return items

    def download_file(self, item_name: str, row_element, current_path: str = "") -> bool:
        """下载单个文件并等待完成"""
        try:
            # 1. 选中该项（点击行）
            try:
                selector = "span[role='checkbox'], div[data-automationid='DetailsRowCheck']"
                checkbox = row_element.find_element(By.CSS_SELECTOR, selector)
                checkbox.click()
            except:
                row_element.click()
            
            time.sleep(1)
            
            # 2. 点击工具栏的“下载”按钮
            download_btn_selectors = [
                "button[name='下载']", 
                "button[name='Download']", 
                "button[data-automationid='downloadCommand']"
            ]
            
            download_btn = None
            for selector in download_btn_selectors:
                try:
                    btns = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if btns:
                        # 确保按钮是可见且可点击的
                        if btns[0].is_displayed():
                            download_btn = btns[0]
                            break
                except:
                    continue
            
            if download_btn:
                download_btn.click()
                logger.info(f"触发下载: {item_name}")
                
                # 3. 监控下载完成
                if self.monitor_download(item_name, current_path):
                    logger.info(f"下载成功: {item_name}")
                    return True
                else:
                    logger.warning(f"下载超时或监控失败: {item_name}")
                    return False
            else:
                logger.error(f"找不到下载按钮: {item_name}")
                return False
                
        except Exception as e:
            logger.error(f"下载过程出错 {item_name}: {str(e)}")
            return False
        finally:
            # 取消选中
            try:
                row_element.click()
            except:
                pass

    def monitor_download(self, filename: str, target_dir: str = "") -> bool:
        """监控下载进度并行移动文件 (移植自 main.py)"""
        timeout = config.SHAREPOINT_DOWNLOAD_TIMEOUT
        download_path = os.path.abspath(config.SHAREPOINT_DOWNLOAD_DIR)
        start_time = time.time()
        last_size = 0
        stable_count = 0
        
        while time.time() - start_time < timeout:
            try:
                files = os.listdir(download_path)
                # 查找匹配的文件
                matching_files = [f for f in files if filename in f or f.startswith(filename.split('.')[0])]
                
                if matching_files:
                    crdownload_files = [f for f in matching_files if f.endswith('.crdownload')]
                    
                    if crdownload_files:
                        download_file_path = os.path.join(download_path, crdownload_files[0])
                        current_size = os.path.getsize(download_file_path)
                        
                        if current_size == last_size:
                            stable_count += 1
                            if stable_count >= 3:
                                time.sleep(2)
                                if os.path.getsize(download_file_path) == current_size:
                                    final_name = crdownload_files[0].replace('.crdownload', '')
                                    temp_path = os.path.join(download_path, final_name)
                                    os.rename(download_file_path, temp_path)
                                    return self.move_file_to_directory(temp_path, filename, target_dir)
                        else:
                            stable_count = 0
                            last_size = current_size
                    else:
                        complete_files = [f for f in matching_files if not f.endswith('.crdownload') and not f.endswith('.tmp')]
                        if complete_files:
                            complete_file_path = os.path.join(download_path, complete_files[0])
                            return self.move_file_to_directory(complete_file_path, filename, target_dir)
                
                time.sleep(3)
            except Exception as e:
                logger.warning(f"监控异常: {str(e)}")
                time.sleep(3)
        return False

    def move_file_to_directory(self, source_path: str, filename: str, target_dir: str) -> bool:
        """移动文件到嵌套目录 (移植自 main.py)"""
        try:
            root_download_dir = config.SHAREPOINT_DOWNLOAD_DIR
            if target_dir:
                dest_dir = os.path.join(root_download_dir, target_dir)
                os.makedirs(dest_dir, exist_ok=True)
                dest_path = os.path.join(dest_dir, filename)
            else:
                dest_path = os.path.join(root_download_dir, filename)
            
            if os.path.exists(dest_path):
                if os.path.getsize(source_path) == os.path.getsize(dest_path):
                    os.remove(source_path)
                    return True
                else:
                    name, ext = os.path.splitext(filename)
                    dest_path = os.path.join(os.path.dirname(dest_path), f"{name}_{int(time.time())}{ext}")
            
            os.rename(source_path, dest_path)
            return True
        except Exception as e:
            logger.error(f"移动文件失败: {str(e)}")
            return False

    def traverse_and_download(self, current_path: str = ""):
        """递归遍历下载"""
        logger.info(f"正在处理目录: {current_path if current_path else '根目录'}")
        
        # 建立本地目录
        local_dir = os.path.join(config.SHAREPOINT_DOWNLOAD_DIR, current_path)
        os.makedirs(local_dir, exist_ok=True)
        
        self.scroll_to_load_all_files()
        items = self.get_items()
        
        # 为了避免 StaleElementReferenceException，先存下信息，每次操作后可能需要重刷
        folders = [it['name'] for it in items if it['is_folder']]
        files = [it['name'] for it in items if not it['is_folder']]
        
        logger.info(f"当前目录发现 {len(folders)} 个文件夹, {len(files)} 个文件")
        
        # 处理文件
        for file_name in files:
            local_file_path = os.path.join(local_dir, file_name)
            if os.path.exists(local_file_path):
                logger.info(f"文件已存在，跳过: {file_name}")
                download_stats['existing_files'] += 1
                continue
                
            # 重新获取元素（因为操作完一个后 DOM 可能会变）
            current_items = self.get_items()
            target_row = next((it['element'] for it in current_items if it['name'] == file_name), None)
            
            if target_row:
                if self.download_file(file_name, target_row, current_path):
                    download_stats['downloaded'] += 1
                else:
                    download_stats['failed'] += 1
            else:
                logger.warning(f"无法重新定位文件行: {file_name}")

        # 处理文件夹
        for folder_name in folders:
            # 重新获取元素进入文件夹
            current_items = self.get_items()
            target_row = next((it['element'] for it in current_items if it['name'] == folder_name), None)
            
            if target_row:
                try:
                    # 点击文件夹名称进入
                    name_btn = target_row.find_element(By.CSS_SELECTOR, "button[data-automationid='FieldRenderer-name']")
                    name_btn.click()
                    time.sleep(3)
                    self.wait_for_page_load()
                    
                    # 递归
                    self.traverse_and_download(os.path.join(current_path, folder_name))
                    
                    # 返回上一级（点击面包屑或后退）
                    # 简单处理：点击倒数第二个面包屑
                    breadcrumbs = self.driver.find_elements(By.CSS_SELECTOR, ".ms-Breadcrumb-itemLink, [data-automationid='Breadcrumb']")
                    if len(breadcrumbs) >= 2:
                        breadcrumbs[-2].click()
                        time.sleep(3)
                        self.wait_for_page_load()
                    else:
                        self.driver.back()
                        time.sleep(3)
                        self.wait_for_page_load()
                        
                except Exception as e:
                    logger.error(f"进入文件夹 {folder_name} 出错: {str(e)}")
            else:
                logger.warning(f"无法重新定位文件夹: {folder_name}")

    def close(self):
        if self.driver:
            self.driver.quit()

def main():
    downloader = SharePointDownloader()
    try:
        if downloader.access_share_link(config.SHAREPOINT_URL):
            downloader.traverse_and_download()
            logger.info("下载流程结束")
            logger.info(f"统计信息: {download_stats}")
    finally:
        downloader.close()

if __name__ == "__main__":
    main()
