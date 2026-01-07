"""
SharePoint 自动化下载工具
使用 Selenium ChromeDriver 自动化访问 SharePoint 共享文件夹并下载文件
"""

import os
import time
import logging
import json
import glob
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Set
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
    'failed': 0,
    'retried': 0
}


class DownloadState:
    """下载状态管理，支持断点续传"""
    
    def __init__(self, download_dir: str):
        self.download_dir = os.path.abspath(download_dir)
        self.state_file = os.path.join(self.download_dir, '.download_state.json')
        self.failed_files: List[Dict] = []  # 失败的文件列表 [{name, path, retries}]
        self.load_state()
    
    def load_state(self):
        """加载之前的下载状态"""
        if os.path.exists(self.state_file):
            try:
                with open(self.state_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.failed_files = data.get('failed_files', [])
                    logger.info(f"加载下载状态: {len(self.failed_files)} 个待重试文件")
            except Exception as e:
                logger.warning(f"加载状态文件失败: {e}")
                self.failed_files = []
    
    def save_state(self):
        """保存下载状态"""
        try:
            os.makedirs(self.download_dir, exist_ok=True)
            with open(self.state_file, 'w', encoding='utf-8') as f:
                json.dump({
                    'failed_files': self.failed_files,
                    'last_updated': datetime.now().isoformat()
                }, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.warning(f"保存状态文件失败: {e}")
    
    def add_failed(self, filename: str, path: str):
        """添加失败的文件"""
        # 检查是否已存在
        for item in self.failed_files:
            if item['name'] == filename and item['path'] == path:
                item['retries'] = item.get('retries', 0) + 1
                self.save_state()
                return
        self.failed_files.append({'name': filename, 'path': path, 'retries': 1})
        self.save_state()
    
    def mark_success(self, filename: str, path: str):
        """标记文件下载成功，从失败列表中移除"""
        self.failed_files = [f for f in self.failed_files 
                            if not (f['name'] == filename and f['path'] == path)]
        self.save_state()
    
    def get_failed_files(self) -> List[Dict]:
        """获取需要重试的文件列表"""
        max_retries = getattr(config, 'SHAREPOINT_MAX_RETRIES', 10)
        return [f for f in self.failed_files if f.get('retries', 0) < max_retries]
    
    def clear_state(self):
        """清除状态"""
        self.failed_files = []
        if os.path.exists(self.state_file):
            os.remove(self.state_file)
    
    def cleanup_incomplete_downloads(self):
        """清理不完整的下载文件（.crdownload, .tmp）"""
        patterns = [
            os.path.join(self.download_dir, '**', '*.crdownload'),
            os.path.join(self.download_dir, '**', '*.tmp'),
        ]
        cleaned = 0
        for pattern in patterns:
            for filepath in glob.glob(pattern, recursive=True):
                try:
                    os.remove(filepath)
                    logger.info(f"已清理不完整文件: {filepath}")
                    cleaned += 1
                except Exception as e:
                    logger.warning(f"清理文件失败 {filepath}: {e}")
        if cleaned > 0:
            logger.info(f"共清理 {cleaned} 个不完整下载文件")
        return cleaned

class SharePointDownloader:
    def __init__(self):
        # 初始化下载状态管理（在driver之前，以便清理文件）
        self.download_state = DownloadState(config.SHAREPOINT_DOWNLOAD_DIR)
        # 清理上次遗留的不完整下载
        self.download_state.cleanup_incomplete_downloads()
        
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
        
        # 首先尝试使用缓存的 ChromeDriver
        driver_path = self._find_cached_chromedriver()
        
        if driver_path:
            logger.info(f"使用缓存的 ChromeDriver: {driver_path}")
            service = Service(driver_path)
        else:
            # 如果没有缓存，则下载
            logger.info("未找到缓存的 ChromeDriver，尝试下载...")
            try:
                service = Service(ChromeDriverManager().install())
            except Exception as e:
                logger.error(f"下载 ChromeDriver 失败: {e}")
                raise RuntimeError("无法获取 ChromeDriver。请检查网络连接，或手动下载 ChromeDriver 并放置到 PATH 中。") from e
        
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_page_load_timeout(config.SHAREPOINT_PAGE_LOAD_TIMEOUT)
        
        logger.info(f"SharePoint ChromeDriver已初始化，下载路径: {download_path}")
        return driver
    
    def _find_cached_chromedriver(self) -> Optional[str]:
        """在 webdriver_manager 缓存目录中查找已存在的 ChromeDriver"""
        import glob
        
        # webdriver_manager 默认缓存目录
        home = os.path.expanduser("~")
        wdm_cache_dir = os.path.join(home, ".wdm", "drivers", "chromedriver")
        
        if not os.path.exists(wdm_cache_dir):
            return None
        
        # 搜索所有 chromedriver.exe 文件
        patterns = [
            os.path.join(wdm_cache_dir, "**", "chromedriver.exe"),  # Windows
            os.path.join(wdm_cache_dir, "**", "chromedriver"),      # Linux/Mac
        ]
        
        driver_files = []
        for pattern in patterns:
            driver_files.extend(glob.glob(pattern, recursive=True))
        
        if not driver_files:
            return None
        
        # 返回最新的（按修改时间排序）
        driver_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        return driver_files[0]

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
            # 等待表格出现
            try:
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='grid'], .ms-List-page")))
            except TimeoutException:
                logger.warning("等待列表容器超时")
                return []

            # SharePoint 常见的行选择器
            # 尝试多种可能的行定位方式
            row_selectors = [
                "div[role='row'][data-automationid^='row-']",
                "div[role='row'][data-automationid='DetailsRow']",
                "div[role='row']",
                ".ms-List-cell"
            ]
            
            rows = []
            for selector in row_selectors:
                rows = self.driver.find_elements(By.CSS_SELECTOR, selector)
                if rows:
                    logger.debug(f"找到 {len(rows)} 行，使用选择器: {selector}")
                    break
            
            for row in rows:
                try:
                    # 获取名称 - 尝试多个属性
                    name = None
                    name_selectors = [
                        "button[data-automationid='FieldRenderer-name']",
                        "[data-automationid='field-LinkFilename']",
                        ".ms-DetailsRow-cell[data-automationid='name']",
                        "a[data-automationid='nameField']"
                    ]
                    
                    for sel in name_selectors:
                        try:
                            name_el = row.find_element(By.CSS_SELECTOR, sel)
                            name = name_el.text.strip()
                            if name: break
                        except:
                            continue
                    
                    if not name:
                        # 兜底：如果没找到具体元素，尝试获取整行文本的第一部分
                        text = row.text.strip()
                        if text:
                            name = text.split('\n')[0].strip()
                    
                    if not name or name == ".." or name == "Nome" or name == "Name":
                        continue
                    
                    # 检查是否为标题行
                    is_header = False
                    try:
                        role = row.get_attribute("role")
                        if role == "columnheader": is_header = True
                        if row.find_elements(By.CSS_SELECTOR, "[role='columnheader']"): is_header = True
                    except:
                        pass
                    
                    if is_header:
                        continue
                    
                    # 判断类型
                    is_folder = False
                    # 1. 检查图标
                    try:
                        icon = row.find_element(By.CSS_SELECTOR, "[data-automationid='field-DocIcon'] i, [data-automationid='field-DocIcon'] img, i[data-icon-name='FabricFolder'], i[data-icon-name='FolderInverse']")
                        aria_label = icon.get_attribute("aria-label") or ""
                        # 支持多语言：folder (英), pasta (葡/意), dossier (法), ordner (德) 等
                        folder_keywords = ["folder", "pasta", "dossier", "ordner", "文件夹", "文件夹", "目录"]
                        if any(kw in aria_label.lower() for kw in folder_keywords):
                            is_folder = True
                    except:
                        pass
                    
                    # 2. 检查是否有 data-automationid="folderIcon"
                    if not is_folder:
                        try:
                            row.find_element(By.CSS_SELECTOR, "[data-automationid='folderIcon']")
                            is_folder = True
                        except:
                            pass

                    # 3. 检查描述文字中是否包含 "items" 或 "个项目" (文件夹通常显示子项数量)
                    if not is_folder:
                        row_text = row.text.lower()
                        if "items" in row_text or "itens" in row_text or "个项目" in row_text:
                            # 排除掉文件名本身包含这些词的情况，这只是一个弱辅助判断
                            pass

                    items.append({
                        'name': name,
                        'is_folder': is_folder,
                        'element': row
                    })
                    logger.debug(f"检测到: {name} (文件夹: {is_folder})")
                except Exception as row_error:
                    continue
        except Exception as e:
            logger.error(f"获取列表失败: {str(e)}")
        
        return items

    def download_file(self, item_name: str, row_element, current_path: str = "") -> bool:
        """下载单个文件并等待完成"""
        try:
            # 确保元素可见
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row_element)
            time.sleep(1)

            # 1. 选中该项
            selected = False
            # 优先尝试点击复选框
            selection_selectors = [
                "input[data-automationid='selection-checkbox']",
                "span[role='checkbox']",
                "div[data-automationid='DetailsRowCheck']",
                "[aria-label*='Selecionar']",
                "[aria-label*='Select']"
            ]
            
            for sel in selection_selectors:
                try:
                    checkbox = row_element.find_element(By.CSS_SELECTOR, sel)
                    if checkbox.is_displayed():
                        # 使用 JS 点击以避免 interactability 问题
                        self.driver.execute_script("arguments[0].click();", checkbox)
                        selected = True
                        break
                except:
                    continue
            
            if not selected:
                # 兜底：直接点击行
                self.driver.execute_script("arguments[0].click();", row_element)
            
            time.sleep(1.5)
            
            # 2. 点击工具栏的“下载”按钮
            download_btn_selectors = [
                "button[data-automationid='downloadCommand']",
                "button[name='下载']", 
                "button[name='Download']",
                "button[name='Baixar']",
                ".ms-Button--primary",
                "i[data-icon-name='Download']"
            ]
            
            download_btn = None
            for selector in download_btn_selectors:
                try:
                    # 在整个文档中找按钮，而不仅仅是在行内
                    btns = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for btn in btns:
                        if btn.is_displayed() and btn.is_enabled():
                            download_btn = btn
                            break
                    if download_btn: break
                except:
                    continue
            
            if download_btn:
                self.driver.execute_script("arguments[0].click();", download_btn)
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
                # 尝试右键菜单作为兜底
                try:
                    actions = ActionChains(self.driver)
                    actions.context_click(row_element).perform()
                    time.sleep(1)
                    # 寻找右键菜单里的下载选项
                    menu_download = self.driver.find_element(By.CSS_SELECTOR, "button[data-automationid='downloadCommand']")
                    menu_download.click()
                    return self.monitor_download(item_name, current_path)
                except:
                    pass
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
        """监控下载进度并移动文件"""
        timeout = config.SHAREPOINT_DOWNLOAD_TIMEOUT
        download_path = os.path.abspath(config.SHAREPOINT_DOWNLOAD_DIR)
        start_time = time.time()
        last_size = 0
        stable_count = 0
        
        # 获取期望的文件扩展名
        expected_ext = os.path.splitext(filename)[1].lower()
        
        while time.time() - start_time < timeout:
            try:
                files = os.listdir(download_path)
                
                # 精确匹配：完整文件名，或者文件名+.crdownload
                exact_match = filename in files
                crdownload_match = f"{filename}.crdownload" in files
                
                # 如果有精确匹配的正在下载的文件
                if crdownload_match:
                    download_file_path = os.path.join(download_path, f"{filename}.crdownload")
                    current_size = os.path.getsize(download_file_path)
                    
                    if current_size == last_size:
                        stable_count += 1
                        if stable_count >= 3:
                            time.sleep(2)
                            if os.path.getsize(download_file_path) == current_size:
                                # 下载完成，重命名去掉 .crdownload
                                final_path = os.path.join(download_path, filename)
                                os.rename(download_file_path, final_path)
                                return self.move_file_to_directory(final_path, filename, target_dir)
                    else:
                        stable_count = 0
                        last_size = current_size
                elif exact_match:
                    # 文件已经下载完成
                    complete_file_path = os.path.join(download_path, filename)
                    return self.move_file_to_directory(complete_file_path, filename, target_dir)
                else:
                    # 尝试模糊匹配（处理 Chrome 可能添加序号的情况，如 file(1).txt）
                    base_name, ext = os.path.splitext(filename)
                    for f in files:
                        # 检查是否是同名文件的变体（带序号）
                        if f.startswith(base_name) and f.endswith(ext) and not f.endswith('.crdownload') and not f.endswith('.tmp'):
                            complete_file_path = os.path.join(download_path, f)
                            # 使用实际的文件名，而不是期望的文件名
                            return self.move_file_to_directory(complete_file_path, f, target_dir)
                        # 检查正在下载的变体
                        if f.startswith(base_name) and f.endswith(f"{ext}.crdownload"):
                            download_file_path = os.path.join(download_path, f)
                            current_size = os.path.getsize(download_file_path)
                            if current_size == last_size:
                                stable_count += 1
                                if stable_count >= 3:
                                    time.sleep(2)
                                    if os.path.getsize(download_file_path) == current_size:
                                        final_name = f.replace('.crdownload', '')
                                        final_path = os.path.join(download_path, final_name)
                                        os.rename(download_file_path, final_path)
                                        return self.move_file_to_directory(final_path, final_name, target_dir)
                            else:
                                stable_count = 0
                                last_size = current_size
                            break
                
                time.sleep(3)
            except Exception as e:
                logger.warning(f"监控异常: {str(e)}")
                time.sleep(3)
        return False

    def move_file_to_directory(self, source_path: str, filename: str, target_dir: str) -> bool:
        """移动文件到嵌套目录 (移植自 main.py)"""
        try:
            root_download_dir = os.path.abspath(config.SHAREPOINT_DOWNLOAD_DIR)
            if target_dir:
                dest_dir = os.path.join(root_download_dir, target_dir)
                os.makedirs(dest_dir, exist_ok=True)
                dest_path = os.path.join(dest_dir, filename)
            else:
                dest_path = os.path.join(root_download_dir, filename)
            
            # 规范化路径以便比较
            source_abs = os.path.abspath(source_path)
            dest_abs = os.path.abspath(dest_path)
            
            # 如果源文件和目标文件是同一个文件，直接返回成功
            if source_abs == dest_abs:
                logger.info(f"文件已在目标位置: {filename}")
                return True
            
            if os.path.exists(dest_path):
                if os.path.getsize(source_path) == os.path.getsize(dest_path):
                    os.remove(source_path)
                    logger.info(f"目标已存在相同文件，删除临时文件: {source_path}")
                    return True
                else:
                    name, ext = os.path.splitext(filename)
                    dest_path = os.path.join(os.path.dirname(dest_path), f"{name}_{int(time.time())}{ext}")
            
            os.rename(source_path, dest_path)
            logger.info(f"文件已移动到: {dest_path}")
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
                # 如果之前在失败列表中，移除
                self.download_state.mark_success(file_name, current_path)
                continue
                
            # 重新获取元素（因为操作完一个后 DOM 可能会变）
            current_items = self.get_items()
            target_row = next((it['element'] for it in current_items if it['name'] == file_name), None)
            
            if target_row:
                if self.download_file(file_name, target_row, current_path):
                    download_stats['downloaded'] += 1
                    self.download_state.mark_success(file_name, current_path)
                else:
                    download_stats['failed'] += 1
                    self.download_state.add_failed(file_name, current_path)
                    logger.warning(f"文件下载失败，已加入重试队列: {file_name}")
            else:
                logger.warning(f"无法重新定位文件行: {file_name}")
                self.download_state.add_failed(file_name, current_path)

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

    def retry_failed_downloads(self, share_url: str) -> int:
        """重试之前失败的下载"""
        failed_files = self.download_state.get_failed_files()
        if not failed_files:
            return 0
        
        logger.info(f"开始重试 {len(failed_files)} 个失败的文件...")
        retried_count = 0
        
        # 按目录分组
        by_path = {}
        for f in failed_files:
            path = f['path']
            if path not in by_path:
                by_path[path] = []
            by_path[path].append(f['name'])
        
        for current_path, file_names in by_path.items():
            # 导航到对应目录
            if current_path:
                # 需要重新访问链接并导航到目录
                logger.info(f"导航到目录: {current_path}")
                self.driver.get(share_url)
                self.wait_for_page_load()
                
                # 逐级进入目录
                path_parts = current_path.split(os.sep)
                for part in path_parts:
                    self.scroll_to_load_all_files()
                    items = self.get_items()
                    folder_row = next((it['element'] for it in items if it['name'] == part and it['is_folder']), None)
                    if folder_row:
                        try:
                            name_btn = folder_row.find_element(By.CSS_SELECTOR, "button[data-automationid='FieldRenderer-name']")
                            name_btn.click()
                            time.sleep(3)
                            self.wait_for_page_load()
                        except Exception as e:
                            logger.error(f"无法进入目录 {part}: {e}")
                            break
                    else:
                        logger.error(f"找不到目录: {part}")
                        break
            else:
                # 根目录，重新访问
                self.driver.get(share_url)
                self.wait_for_page_load()
            
            # 重试该目录下的文件
            self.scroll_to_load_all_files()
            local_dir = os.path.join(config.SHAREPOINT_DOWNLOAD_DIR, current_path)
            
            for file_name in file_names:
                local_file_path = os.path.join(local_dir, file_name)
                if os.path.exists(local_file_path):
                    logger.info(f"重试文件已存在，跳过: {file_name}")
                    self.download_state.mark_success(file_name, current_path)
                    continue
                
                current_items = self.get_items()
                target_row = next((it['element'] for it in current_items if it['name'] == file_name), None)
                
                if target_row:
                    logger.info(f"重试下载: {file_name}")
                    if self.download_file(file_name, target_row, current_path):
                        download_stats['retried'] += 1
                        self.download_state.mark_success(file_name, current_path)
                        retried_count += 1
                    else:
                        self.download_state.add_failed(file_name, current_path)
                else:
                    logger.warning(f"重试时无法定位文件: {file_name}")
        
        return retried_count

    def close(self):
        if self.driver:
            self.driver.quit()

def main():
    downloader = SharePointDownloader()
    try:
        if downloader.access_share_link(config.SHAREPOINT_URL):
            downloader.traverse_and_download()
            
            # 尝试重试失败的文件
            failed_count = len(downloader.download_state.get_failed_files())
            if failed_count > 0:
                logger.info(f"首轮下载完成，{failed_count} 个文件失败，开始重试...")
                retried = downloader.retry_failed_downloads(config.SHAREPOINT_URL)
                logger.info(f"重试完成，成功 {retried} 个文件")
            
            logger.info("下载流程结束")
            logger.info(f"统计信息: {download_stats}")
            
            # 显示仍然失败的文件
            remaining_failed = downloader.download_state.get_failed_files()
            if remaining_failed:
                logger.warning(f"仍有 {len(remaining_failed)} 个文件下载失败:")
                for f in remaining_failed[:10]:  # 只显示前10个
                    logger.warning(f"  - {os.path.join(f['path'], f['name'])} (尝试 {f.get('retries', 0)} 次)")
                if len(remaining_failed) > 10:
                    logger.warning(f"  ... 还有 {len(remaining_failed) - 10} 个文件")
            else:
                # 全部成功，清除状态文件
                downloader.download_state.clear_state()
                logger.info("所有文件下载成功！")
    finally:
        downloader.close()

if __name__ == "__main__":
    main()
