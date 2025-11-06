"""
ownCloud自动化下载工具
使用Selenium ChromeDriver自动化登录、遍历文件夹并逐个下载文件
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
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import config

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('owncloud_download.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 全局变量：存储下载失败的文件
failed_files: List[Tuple[str, str, str]] = []

# 全局变量：下载统计信息
download_stats = {
    'total_files': 0,      # 发现的文件总数
    'existing_files': 0,   # 已存在的文件数
    'downloaded': 0,       # 本轮成功下载的文件数
    'failed': 0            # 本轮失败的文件数
}


def setup_chrome_driver() -> webdriver.Chrome:
    """
    初始化ChromeDriver，配置下载路径和选项
    """
    chrome_options = Options()
    
    # 设置下载路径（绝对路径）
    download_path = os.path.abspath(config.DOWNLOAD_DIR)
    os.makedirs(download_path, exist_ok=True)
    
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # 启用显示界面（headless=False）
    # chrome_options.add_argument("--headless")  # 注释掉，显示界面
    
    # 其他选项
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # 创建Service和Driver
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_page_load_timeout(config.PAGE_LOAD_TIMEOUT)
    
    logger.info(f"ChromeDriver已初始化，下载路径: {download_path}")
    return driver


def login_to_owncloud(driver: webdriver.Chrome) -> bool:
    """
    访问ownCloud分享链接并输入密码
    """
    try:
        logger.info(f"正在访问ownCloud链接: {config.OWNCLOUD_URL}")
        driver.get(config.OWNCLOUD_URL)
        
        # 等待密码输入框出现
        wait = WebDriverWait(driver, config.PAGE_LOAD_TIMEOUT)
        password_input = wait.until(
            EC.presence_of_element_located((By.ID, "password"))
        )
        
        logger.info("找到密码输入框，正在输入密码...")
        password_input.clear()
        password_input.send_keys(config.SHARE_PASSWORD)
        
        # 点击提交按钮
        submit_button = driver.find_element(By.ID, "password-submit")
        submit_button.click()
        
        # 等待页面加载完成（等待文件列表出现或URL变化）
        time.sleep(3)
        wait.until(lambda d: "password" not in d.current_url.lower() or 
                   d.find_elements(By.CLASS_NAME, "filelist") or
                   d.find_elements(By.ID, "fileList"))
        
        logger.info("登录成功！")
        return True
        
    except TimeoutException:
        logger.error("登录超时，密码输入框未出现")
        return False
    except Exception as e:
        logger.error(f"登录过程中发生错误: {str(e)}")
        return False


def wait_for_page_load(driver: webdriver.Chrome, timeout: int = 10) -> bool:
    """
    等待页面JavaScript加载完成
    """
    try:
        # 等待document.readyState为complete
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        # 等待jQuery加载完成（如果使用jQuery）
        try:
            WebDriverWait(driver, timeout).until(
                lambda d: d.execute_script("return typeof jQuery !== 'undefined' && jQuery.active == 0")
            )
        except:
            pass  # 如果没有jQuery，忽略
        
        # 额外等待一点时间确保所有内容加载
        time.sleep(2)
        return True
        
    except TimeoutException:
        logger.warning("页面加载超时，但继续执行")
        return False


def scroll_to_load_all_files(driver: webdriver.Chrome) -> None:
    """
    滚动页面直到所有文件都被加载（无法再滚动为止）
    """
    try:
        logger.debug("开始滚动页面以加载所有文件...")
        
        # 获取文件列表容器（如果有的话）
        file_list_container = None
        try:
            file_list_container = driver.find_element(By.CSS_SELECTOR, "#fileList, .filelist, .files-fileList, tbody, #content-wrapper, .content")
        except:
            pass
        
        # 定义滚动函数
        def get_scroll_height():
            if file_list_container:
                return driver.execute_script("return arguments[0].scrollHeight;", file_list_container)
            else:
                return driver.execute_script("return document.body.scrollHeight;")
        
        def get_scroll_top():
            if file_list_container:
                return driver.execute_script("return arguments[0].scrollTop;", file_list_container)
            else:
                return driver.execute_script("return window.pageYOffset || document.documentElement.scrollTop;")
        
        def scroll_to_bottom():
            if file_list_container:
                driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", file_list_container)
            else:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        
        last_height = get_scroll_height()
        last_scroll_top = get_scroll_top()
        scroll_attempts = 0
        max_scroll_attempts = 200  # 防止无限循环
        no_change_count = 0
        max_no_change = 3  # 连续3次没有变化就认为滚动到底了
        
        while scroll_attempts < max_scroll_attempts:
            # 滚动到页面底部
            scroll_to_bottom()
            
            # 等待新内容加载（ownCloud可能需要时间加载更多文件）
            time.sleep(1.5)
            
            # 获取新的滚动位置和页面高度
            new_height = get_scroll_height()
            new_scroll_top = get_scroll_top()
            
            # 检查是否真的滚动到底了
            max_scroll = new_height - (driver.execute_script("return window.innerHeight;") if not file_list_container else file_list_container.size['height'])
            is_at_bottom = abs(new_scroll_top - max_scroll) < 10  # 允许10像素误差
            
            # 检查是否有新内容加载
            if new_height == last_height and is_at_bottom:
                no_change_count += 1
                if no_change_count >= max_no_change:
                    logger.debug(f"页面已滚动到底部，无法再滚动（连续{no_change_count}次无变化）")
                    break
            else:
                no_change_count = 0  # 有变化，重置计数
                if new_height > last_height:
                    logger.debug(f"页面高度变化: {last_height} -> {new_height}，继续滚动...")
            
            last_height = new_height
            last_scroll_top = new_scroll_top
            scroll_attempts += 1
        
        # 滚动回顶部，确保所有元素都在视图中
        if file_list_container:
            driver.execute_script("arguments[0].scrollTop = 0;", file_list_container)
        else:
            driver.execute_script("window.scrollTo(0, 0);")
        
        time.sleep(0.5)  # 等待滚动完成
        
        logger.info(f"滚动完成，共尝试 {scroll_attempts} 次，最终页面高度: {last_height}")
        
    except Exception as e:
        logger.warning(f"滚动页面时出错: {str(e)}")


def get_file_list(driver: webdriver.Chrome) -> List[Tuple[str, bool, any]]:
    """
    获取当前目录下的文件和文件夹列表
    返回: [(名称, 是否为文件夹, WebElement), ...]
    """
    file_list = []
    
    try:
        wait_for_page_load(driver)
        
        # 先滚动页面，确保所有文件都被加载
        scroll_to_load_all_files(driver)
        
        # ownCloud的文件列表通常在特定容器中
        # 尝试多种可能的选择器
        file_elements = []
        
        # 方法1: 通过data-file属性查找
        try:
            file_elements = driver.find_elements(By.CSS_SELECTOR, "[data-file], [data-type], .file, tr.file, .filelist tbody tr")
        except:
            pass
        
        # 方法2: 如果没有找到，尝试其他选择器
        if not file_elements:
            try:
                file_elements = driver.find_elements(By.CSS_SELECTOR, "tbody tr, .files-fileList tr")
            except:
                pass
        
        # 方法3: 通过包含文件名的元素查找
        if not file_elements:
            try:
                # 查找所有可能的文件行
                file_elements = driver.find_elements(By.XPATH, "//tr[contains(@class, 'file') or contains(@data-type, 'file') or contains(@data-type, 'dir')]")
            except:
                pass
        
        logger.info(f"找到 {len(file_elements)} 个文件/文件夹项")
        
        for element in file_elements:
            try:
                # 获取文件名（可能在不同的子元素中）
                name = None
                name_elements = element.find_elements(By.CSS_SELECTOR, ".name, .filename, td.name, a.name, .file-name")
                if name_elements:
                    name = name_elements[0].text.strip()
                else:
                    # 如果没有找到，尝试获取整个元素的文本
                    name = element.text.strip().split('\n')[0]
                
                if not name:
                    continue
                
                # 跳过特殊项（如".."或空名称）
                if name == ".." or name == "." or not name:
                    continue
                
                # 判断是否为文件夹
                is_folder = is_folder_element(element, name)
                
                file_list.append((name, is_folder, element))
                logger.debug(f"找到: {name} ({'文件夹' if is_folder else '文件'})")
                
            except Exception as e:
                logger.warning(f"解析文件项时出错: {str(e)}")
                continue
        
        return file_list
        
    except Exception as e:
        logger.error(f"获取文件列表时出错: {str(e)}")
        return []


def is_folder_element(element: any, name: str) -> bool:
    """
    判断元素是否为文件夹
    """
    try:
        # 方法1: 检查data-type属性
        data_type = element.get_attribute("data-type")
        if data_type:
            return "dir" in data_type.lower() or "folder" in data_type.lower()
        
        # 方法2: 检查class名称
        class_name = element.get_attribute("class")
        if class_name:
            if "dir" in class_name.lower() or "folder" in class_name.lower():
                return True
            if "file" in class_name.lower() and "dir" not in class_name.lower():
                return False
        
        # 方法3: 检查图标
        icons = element.find_elements(By.CSS_SELECTOR, ".icon, img[src*='folder'], img[src*='directory']")
        if icons:
            icon_src = icons[0].get_attribute("src") or ""
            if "folder" in icon_src.lower() or "directory" in icon_src.lower():
                return True
        
        # 方法4: 检查mimetype（如果存在）
        mimetype = element.get_attribute("data-mimetype")
        if mimetype:
            return "directory" in mimetype.lower() or "folder" in mimetype.lower()
        
        # 方法5: 默认情况下，如果名称没有扩展名，可能是文件夹（不可靠，作为最后手段）
        # 这里我们返回False，让用户或更可靠的检查来决定
        
        return False
        
    except Exception as e:
        logger.warning(f"判断文件夹时出错: {str(e)}")
        return False


def create_local_directory(path: str) -> bool:
    """
    创建本地目录结构
    """
    try:
        full_path = os.path.join(config.DOWNLOAD_DIR, path)
        os.makedirs(full_path, exist_ok=True)
        logger.debug(f"创建目录: {full_path}")
        return True
    except Exception as e:
        logger.error(f"创建目录失败 {path}: {str(e)}")
        return False


def scan_directory(driver: webdriver.Chrome, current_path: str = "") -> None:
    """
    递归遍历所有文件夹并建立目录树，同时收集文件下载信息
    """
    try:
        logger.info(f"正在扫描目录: {current_path if current_path else '根目录'}")
        
        # 创建当前目录
        if current_path:
            create_local_directory(current_path)
        
        # 等待页面加载
        wait_for_page_load(driver)
        time.sleep(2)  # 额外等待确保页面稳定
        
        # 获取文件列表（内部会先滚动页面加载所有文件）
        file_list = get_file_list(driver)
        
        if not file_list:
            logger.warning(f"目录 {current_path} 中没有找到文件或文件夹")
            return
        
        # 分离文件和文件夹，避免stale element reference错误
        files = []
        folders = []
        
        for name, is_folder, element in file_list:
            try:
                if is_folder:
                    folders.append((name, element))
                else:
                    files.append((name, element))
            except Exception as e:
                logger.warning(f"分类项时出错 {name}: {str(e)}")
                continue
        
        # 先处理所有文件，立即下载
        for name, element in files:
            try:
                file_path = os.path.join(current_path, name).replace("\\", "/") if current_path else name
                
                # 统计：发现的文件总数
                download_stats['total_files'] += 1
                
                # 存储元素定位信息（使用XPath）
                try:
                    element_xpath = get_element_xpath(driver, element)
                    file_info = (current_path, name, element_xpath)
                    logger.info(f"发现文件，立即下载: {file_path}")
                except Exception as e:
                    logger.warning(f"无法获取元素定位信息 {name}: {str(e)}")
                    # 如果获取XPath失败，使用文件名作为备用定位方式
                    file_info = (current_path, name, "")
                
                # 检查文件是否已存在
                local_file_path = os.path.join(config.DOWNLOAD_DIR, current_path, name) if current_path else os.path.join(config.DOWNLOAD_DIR, name)
                local_file_path = os.path.normpath(local_file_path)
                
                if os.path.exists(local_file_path) and os.path.isfile(local_file_path):
                    file_size = os.path.getsize(local_file_path)
                    logger.info(f"文件已存在，跳过下载: {file_path} (大小: {file_size} 字节)")
                    download_stats['existing_files'] += 1
                    continue
                
                # 文件不存在，立即下载（当前已在文件所在目录，无需导航）
                # 使用重试机制下载
                logger.info(f"文件不存在，开始下载: {file_path}")
                if download_and_monitor_with_retry(driver, file_info, current_path):
                    logger.info(f"下载完成: {file_path}")
                    download_stats['downloaded'] += 1
                else:
                    # 下载失败，添加到失败列表（稍后重试）
                    logger.error(f"下载失败，已添加到重试列表: {file_path}")
                    failed_files.append(file_info)
                    download_stats['failed'] += 1
                    
            except Exception as e:
                logger.error(f"处理文件 {name} 时出错: {str(e)}")
                continue
        
        # 然后处理所有文件夹
        for name, element in folders:
            try:
                folder_path = os.path.join(current_path, name).replace("\\", "/") if current_path else name
                logger.info(f"进入文件夹: {folder_path}")
                
                # 在点击之前，通过文件名重新定位元素（避免使用可能失效的element引用）
                folder_element = None
                try:
                    # 尝试通过文件名定位文件夹
                    all_elements = driver.find_elements(By.CSS_SELECTOR, "[data-file], .file, tr.file, tbody tr")
                    for elem in all_elements:
                        elem_text = elem.text.strip()
                        if name in elem_text and name == elem_text.split('\n')[0].strip():
                            # 检查是否为文件夹
                            if is_folder_element(elem, name):
                                folder_element = elem
                                break
                except Exception as e:
                    logger.warning(f"重新定位文件夹元素失败 {name}: {str(e)}")
                
                if folder_element:
                    # 点击文件夹进入
                    clickable = None
                    try:
                        # 尝试找到可点击的元素（链接或文件名）
                        clickable = folder_element.find_element(By.CSS_SELECTOR, "a, .name, .filename, td.name")
                    except:
                        try:
                            clickable = folder_element
                        except:
                            pass
                    
                    if clickable:
                        clickable.click()
                        time.sleep(3)  # 等待页面导航
                        wait_for_page_load(driver)
                        
                        # 递归扫描子目录
                        scan_directory(driver, folder_path)
                        
                        # 返回上级目录（通过breadcrumb或返回按钮）
                        navigate_back(driver)
                        time.sleep(2)
                        wait_for_page_load(driver)
                    else:
                        logger.warning(f"无法找到可点击元素进入文件夹: {name}")
                else:
                    logger.warning(f"无法定位文件夹元素: {name}")
                        
            except Exception as e:
                logger.error(f"处理文件夹 {name} 时出错: {str(e)}")
                continue
                
    except Exception as e:
        logger.error(f"扫描目录 {current_path} 时出错: {str(e)}")


def get_element_xpath(driver: webdriver.Chrome, element: any) -> str:
    """
    获取元素的XPath定位
    """
    try:
        return driver.execute_script("""
            function getElementXPath(element) {
                if (element.id !== '') {
                    return '//*[@id="' + element.id + '"]';
                }
                if (element === document.body) {
                    return '/html/body';
                }
                var ix = 0;
                var siblings = element.parentNode.childNodes;
                for (var i = 0; i < siblings.length; i++) {
                    var sibling = siblings[i];
                    if (sibling === element) {
                        return getElementXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';
                    }
                    if (sibling.nodeType === 1 && sibling.tagName === element.tagName) {
                        ix++;
                    }
                }
            }
            return getElementXPath(arguments[0]);
        """, element)
    except:
        # 如果JavaScript方法失败，返回一个基于文本的定位策略
        try:
            name = element.text.strip().split('\n')[0]
            return f"//*[contains(text(), '{name}')]"
        except:
            return ""


def navigate_back(driver: webdriver.Chrome) -> bool:
    """
    导航回上级目录
    """
    try:
        # 方法1: 点击breadcrumb中的上级目录
        breadcrumbs = driver.find_elements(By.CSS_SELECTOR, ".breadcrumb a, .crumb a, nav a")
        if breadcrumbs and len(breadcrumbs) > 1:
            # 点击倒数第二个（返回上一级）
            breadcrumbs[-2].click()
            time.sleep(2)
            return True
        
        # 方法2: 使用浏览器后退按钮
        driver.back()
        time.sleep(2)
        return True
        
    except Exception as e:
        logger.warning(f"返回上级目录失败: {str(e)}")
        # 尝试浏览器后退
        try:
            driver.back()
            time.sleep(2)
            return True
        except:
            return False


def download_and_monitor_with_retry(driver: webdriver.Chrome, file_info: Tuple[str, str, str], current_path: str, max_retries: int = None) -> bool:
    """
    下载文件并监控，包含重试逻辑（最多10次）
    """
    if max_retries is None:
        max_retries = config.MAX_RETRIES
    
    current_path, filename, element_xpath = file_info
    
    for attempt in range(max_retries):
        try:
            if attempt > 0:
                logger.info(f"尝试下载 {filename} (第 {attempt + 1}/{max_retries} 次)")
            
            # 触发下载
            if download_file_in_current_directory(driver, file_info):
                # 监控下载进度
                if monitor_download(driver, filename, current_path):
                    logger.info(f"成功下载: {filename}")
                    return True
                else:
                    logger.warning(f"下载监控失败: {filename}")
            else:
                logger.warning(f"触发下载失败: {filename}")
            
            # 如果失败，等待后重试
            if attempt < max_retries - 1:
                wait_time = config.RETRY_WAIT_TIME * (attempt + 1)  # 递增等待时间
                logger.info(f"等待 {wait_time} 秒后重试...")
                time.sleep(wait_time)
                
        except Exception as e:
            logger.error(f"下载尝试 {attempt + 1} 出错: {str(e)}")
            if attempt < max_retries - 1:
                time.sleep(config.RETRY_WAIT_TIME)
    
    logger.error(f"下载失败，已重试 {max_retries} 次: {filename}")
    return False


def monitor_download(driver: webdriver.Chrome, filename: str, target_dir: str = "", timeout: int = None) -> bool:
    """
    监控Chrome下载状态，下载完成后移动到目标目录
    """
    if timeout is None:
        timeout = config.DOWNLOAD_TIMEOUT
    
    download_path = os.path.abspath(config.DOWNLOAD_DIR)
    start_time = time.time()
    last_size = 0
    stable_count = 0
    
    logger.info(f"开始监控下载: {filename}")
    
    while time.time() - start_time < timeout:
        try:
            # 检查下载目录中的文件
            files = os.listdir(download_path)
            
            # 查找匹配的文件（可能是完整文件名或带.crdownload后缀）
            matching_files = [f for f in files if filename in f or f.startswith(filename.split('.')[0])]
            
            if matching_files:
                # 检查是否有.crdownload文件（Chrome下载中）
                crdownload_files = [f for f in matching_files if f.endswith('.crdownload')]
                
                if crdownload_files:
                    # 仍在下载中，检查文件大小变化
                    download_file_path = os.path.join(download_path, crdownload_files[0])
                    current_size = os.path.getsize(download_file_path)
                    
                    if current_size == last_size:
                        stable_count += 1
                        if stable_count >= 3:  # 连续3次检查大小不变，可能已完成
                            time.sleep(2)  # 再等2秒确认
                            if os.path.getsize(download_file_path) == current_size:
                                # 重命名文件（移除.crdownload后缀）
                                final_name = crdownload_files[0].replace('.crdownload', '')
                                temp_path = os.path.join(download_path, final_name)
                                os.rename(download_file_path, temp_path)
                                
                                # 移动到目标目录
                                if move_file_to_directory(temp_path, filename, target_dir):
                                    logger.info(f"下载完成并移动到目标目录: {filename}")
                                    return True
                                else:
                                    logger.warning(f"下载完成但移动失败: {filename}")
                                    return False
                    else:
                        stable_count = 0
                        last_size = current_size
                        logger.debug(f"下载中: {filename}, 当前大小: {current_size} 字节")
                else:
                    # 没有.crdownload文件，检查是否有完整文件
                    complete_files = [f for f in matching_files if not f.endswith('.crdownload') and not f.endswith('.tmp')]
                    if complete_files:
                        # 找到完整文件，移动到目标目录
                        complete_file_path = os.path.join(download_path, complete_files[0])
                        if move_file_to_directory(complete_file_path, filename, target_dir):
                            logger.info(f"下载完成并移动到目标目录: {filename}")
                            return True
                        else:
                            logger.warning(f"下载完成但移动失败: {filename}")
                            return False
            
            time.sleep(5)  # 每5秒检查一次
            
        except Exception as e:
            logger.warning(f"监控下载时出错: {str(e)}")
            time.sleep(5)
    
    logger.error(f"下载超时: {filename}")
    return False


def move_file_to_directory(source_path: str, filename: str, target_dir: str) -> bool:
    """
    将下载的文件移动到目标目录
    """
    try:
        # 构建目标路径
        if target_dir:
            target_path = os.path.join(config.DOWNLOAD_DIR, target_dir)
            os.makedirs(target_path, exist_ok=True)
            dest_path = os.path.join(target_path, filename)
        else:
            dest_path = os.path.join(config.DOWNLOAD_DIR, filename)
        
        # 如果目标文件已存在，先删除或重命名
        if os.path.exists(dest_path):
            # 检查文件是否相同（通过大小和修改时间）
            if os.path.getsize(source_path) == os.path.getsize(dest_path):
                logger.debug(f"目标文件已存在且相同，删除源文件: {filename}")
                os.remove(source_path)
                return True
            else:
                # 文件不同，添加时间戳后缀
                name, ext = os.path.splitext(filename)
                timestamp = int(time.time())
                dest_path = os.path.join(os.path.dirname(dest_path), f"{name}_{timestamp}{ext}")
        
        # 移动文件
        os.rename(source_path, dest_path)
        logger.debug(f"文件已移动到: {dest_path}")
        return True
        
    except Exception as e:
        logger.error(f"移动文件失败 {filename} 到 {target_dir}: {str(e)}")
        return False


def download_file_in_current_directory(driver: webdriver.Chrome, file_info: Tuple[str, str, str]) -> bool:
    """
    在当前目录中选中文件并下载（不需要导航，因为已经在文件所在目录）
    """
    current_path, filename, element_xpath = file_info
    
    try:
        # 先取消所有已选中的项，确保只选中当前文件
        try:
            # 查找所有已选中的复选框并取消选中
            selected_checkboxes = driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']:checked, .select-checkbox:checked, .checkbox:checked")
            for cb in selected_checkboxes:
                cb.click()
                time.sleep(0.2)
        except:
            pass
        
        # 等待页面稳定
        time.sleep(0.5)
        
        # 定位文件元素（通过文件名精确匹配）
        file_element = None
        try:
            # 方法1: 先尝试通过XPath定位（如果有效）
            if element_xpath and element_xpath.startswith("//"):
                try:
                    file_element = driver.find_element(By.XPATH, element_xpath)
                    # 验证是否是正确的文件（通过文本内容）
                    if filename not in file_element.text:
                        file_element = None
                except:
                    file_element = None
            
            # 方法2: 通过文件名精确匹配查找文件行
            if not file_element:
                all_elements = driver.find_elements(By.CSS_SELECTOR, "[data-file], .file, tr.file, tbody tr")
                for elem in all_elements:
                    try:
                        elem_text = elem.text.strip()
                        # 获取第一行文本（文件名）
                        first_line = elem_text.split('\n')[0].strip() if '\n' in elem_text else elem_text.strip()
                        # 精确匹配文件名
                        if first_line == filename:
                            # 验证不是文件夹
                            if not is_folder_element(elem, filename):
                                file_element = elem
                                break
                    except:
                        continue
            
            # 方法3: 通过XPath查找包含文件名的行
            if not file_element:
                try:
                    candidates = driver.find_elements(By.XPATH, f"//tr[contains(., '{filename}')]")
                    for cand in candidates:
                        if filename in cand.text and not is_folder_element(cand, filename):
                            file_element = cand
                            break
                except:
                    pass
                    
        except Exception as e:
            logger.warning(f"定位文件元素时出错: {str(e)}")
        
        if not file_element:
            logger.error(f"无法定位文件元素: {filename}")
            return False
        
        # 选中文件：先鼠标悬停在文件行左侧显示复选框，然后选中
        checkbox_found = False
        try:
            # 先滚动到元素可见
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", file_element)
            time.sleep(0.3)
            
            # 获取元素的位置和大小
            location = file_element.location
            size = file_element.size
            
            # 方法1: 尝试找到文件行的第一个单元格（复选框通常在这里）
            first_cell = None
            try:
                first_cell = file_element.find_element(By.CSS_SELECTOR, "td:first-child, th:first-child")
            except:
                # 如果找不到，尝试其他方式定位左侧区域
                pass
            
            # 鼠标移动到文件行的左侧区域（复选框出现的位置）
            actions = ActionChains(driver)
            if first_cell:
                # 移动到第一列的左侧边缘
                actions.move_to_element_with_offset(first_cell, -first_cell.size['width']//4, first_cell.size['height']//2)
            else:
                # 移动到文件行的左侧边缘
                actions.move_to_element_with_offset(file_element, -size['width']//4, size['height']//2)
            
            actions.perform()
            time.sleep(1.0)  # 等待复选框出现（ownCloud需要时间显示复选框）
            
            # 在多个位置查找复选框：文件行本身、第一列、文件行的父元素等
            checkbox = None
            search_locations = [
                file_element,  # 文件行本身
                driver,  # 整个页面（复选框可能不在file_element内部）
            ]
            
            if first_cell:
                search_locations.insert(0, first_cell)  # 第一列优先
            
            # 尝试从文件行向上查找父元素
            try:
                parent = file_element.find_element(By.XPATH, "./..")
                search_locations.insert(0, parent)
            except:
                pass
            
            for search_location in search_locations:
                try:
                    # 查找复选框的多种可能选择器
                    selectors = [
                        "input[type='checkbox']",
                        ".select-checkbox",
                        ".checkbox",
                        "[type='checkbox']",
                        "input.checkbox",
                        ".select-checkbox input",
                        "td input[type='checkbox']",
                    ]
                    
                    for selector in selectors:
                        try:
                            if search_location == driver:
                                checkbox = driver.find_element(By.CSS_SELECTOR, selector)
                            else:
                                checkbox = search_location.find_element(By.CSS_SELECTOR, selector)
                            
                            # 验证复选框是否与当前文件相关（通过检查位置）
                            checkbox_loc = checkbox.location
                            # 复选框应该在与文件行相同或接近的Y坐标
                            if abs(checkbox_loc['y'] - location['y']) < 100:
                                logger.debug(f"找到复选框（选择器: {selector}）")
                                break
                        except:
                            continue
                    
                    if checkbox:
                        break
                except:
                    continue
            
            # 如果找到了复选框，点击它
            if checkbox:
                try:
                    # 确保复选框可见
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
                    time.sleep(0.3)
                    
                    # 使用JavaScript点击，更可靠
                    driver.execute_script("arguments[0].click();", checkbox)
                    checkbox_found = True
                    logger.info(f"成功勾选复选框选中文件: {filename}")
                    time.sleep(0.5)  # 等待选中状态生效
                except Exception as e:
                    logger.warning(f"点击复选框失败: {str(e)}")
                    # 备用：使用Selenium点击
                    try:
                        checkbox.click()
                        checkbox_found = True
                        logger.info(f"通过复选框选中文件（备用方法）: {filename}")
                    except:
                        pass
            
            # 如果仍然找不到复选框，尝试直接点击左侧区域
            if not checkbox_found:
                try:
                    logger.debug("复选框未找到，尝试点击左侧区域")
                    actions = ActionChains(driver)
                    if first_cell:
                        actions.move_to_element(first_cell)
                    else:
                        actions.move_to_element_with_offset(file_element, -size['width']//4, size['height']//2)
                    actions.click()
                    actions.perform()
                    time.sleep(0.8)
                    
                    # 点击后再次尝试查找复选框
                    for search_location in search_locations:
                        try:
                            checkbox = search_location.find_element(By.CSS_SELECTOR, "input[type='checkbox']")
                            if checkbox and not checkbox.is_selected():
                                checkbox.click()
                                checkbox_found = True
                                logger.info(f"点击左侧区域后找到并选中复选框: {filename}")
                                break
                        except:
                            continue
                except Exception as e:
                    logger.warning(f"点击左侧区域失败: {str(e)}")
        
        except Exception as e:
            logger.warning(f"鼠标悬停操作失败 {filename}: {str(e)}")
        
        if not checkbox_found:
            logger.error(f"无法找到并勾选复选框: {filename}")
            return False
        
        # 验证文件是否真的被选中
        time.sleep(0.5)
        try:
            # 检查文件行是否有选中状态（通过class或其他属性）
            file_class = file_element.get_attribute("class")
            if "selected" in file_class.lower() or "highlighted" in file_class.lower():
                logger.debug(f"文件已确认选中: {filename}")
        except:
            pass
        
        time.sleep(1)  # 等待选中状态生效
        
        # 获取下载链接并交给Chrome下载器（不直接点击链接）
        # 由于是中文版，应该查找"下载"按钮，或者通过href属性查找包含/download?path=的链接
        download_url = None
        
        # 方法1: 通过href属性查找下载链接（最可靠的方法）
        # ownCloud的下载链接格式: /index.php/s/[token]/download?path=[path]&files=[filename]
        try:
            # 等待一下让下载按钮出现
            time.sleep(0.5)
            
            # 构建预期的下载链接（包含文件名）
            # 文件名可能包含空格，在URL中会被编码为%20或+
            expected_files = filename
            # 处理文件名中的特殊字符：空格可能被编码为%20或+
            expected_files_encoded_space = expected_files.replace(' ', '%20')
            expected_files_encoded_plus = expected_files.replace(' ', '+')
            
            # 方法1: 先尝试精确匹配（包含文件名，处理空格编码）
            download_links = driver.find_elements(By.XPATH, 
                f"//a[contains(@href, '/download?path=') and (contains(@href, '{expected_files}') or contains(@href, '{expected_files_encoded_space}') or contains(@href, '{expected_files_encoded_plus}'))]")
            
            if not download_links:
                # 方法2: 如果精确匹配失败，查找所有包含download?path=的链接，然后手动过滤
                download_links = driver.find_elements(By.XPATH, "//a[contains(@href, '/download?path=')]")
                # 过滤出包含文件名的链接（处理各种编码）
                filtered_links = []
                for link in download_links:
                    href = link.get_attribute("href")
                    if href:
                        # 检查href是否包含文件名（原始、URL编码空格、+编码空格）
                        if (expected_files in href or 
                            expected_files_encoded_space in href or 
                            expected_files_encoded_plus in href):
                            filtered_links.append(link)
                download_links = filtered_links
            
            # 方法3: 如果还是找不到，尝试匹配文件名的主要部分（不含扩展名或部分名称）
            if not download_links:
                # 获取文件名的主要部分（去除扩展名，处理空格）
                name_parts = filename.rsplit('.', 1)[0]  # 文件名不含扩展名
                if ' ' in name_parts:
                    # 如果包含空格，尝试匹配名称的主要单词
                    main_words = [w for w in name_parts.split() if len(w) > 3]  # 长度大于3的词
                    if main_words:
                        for word in main_words:
                            download_links = driver.find_elements(By.XPATH, f"//a[contains(@href, '/download?path=') and contains(@href, '{word}')]")
                            if download_links:
                                break
            
            # 排除header区域的链接
            header_elements = driver.find_elements(By.CSS_SELECTOR, "header, .header, #header")
            header_areas = []
            for header in header_elements:
                loc = header.location
                sz = header.size
                header_areas.append({
                    'top': loc['y'],
                    'bottom': loc['y'] + sz['height'],
                    'left': loc['x'],
                    'right': loc['x'] + sz['width']
                })
            
            for link in download_links:
                link_loc = link.location
                # 检查链接是否在header区域
                in_header = False
                for header_area in header_areas:
                    if (header_area['top'] <= link_loc['y'] <= header_area['bottom'] and
                        header_area['left'] <= link_loc['x'] <= header_area['right']):
                        in_header = True
                        break
                
                # 如果不在header中，且包含文件名，就是我们要找的链接
                if not in_header:
                    href = link.get_attribute("href")
                    if href:
                        # 检查href是否包含文件名（处理空格编码）
                        if (expected_files in href or 
                            expected_files_encoded_space in href or 
                            expected_files_encoded_plus in href):
                            download_url = href
                            logger.info(f"找到下载链接: {download_url}")
                            break
                        # 额外检查：URL解码后是否包含文件名
                        try:
                            from urllib.parse import unquote
                            decoded_href = unquote(href)
                            if expected_files in decoded_href:
                                download_url = href
                                logger.info(f"找到下载链接（URL解码匹配）: {download_url}")
                                break
                        except:
                            pass
            
        except Exception as e:
            logger.warning(f"通过href查找下载链接失败: {str(e)}")
        
        # 方法2: 查找中文"下载"按钮/链接（在文件列表区域，不在header中）
        if not download_url:
            try:
                # 查找包含"下载"文本的按钮或链接
                download_elements = WebDriverWait(driver, 5).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//button[contains(text(), '下载')] | //a[contains(text(), '下载')] | //span[contains(text(), '下载')]"))
                )
                
                # 排除header区域
                header_elements = driver.find_elements(By.CSS_SELECTOR, "header, .header, #header")
                header_areas = []
                for header in header_elements:
                    loc = header.location
                    sz = header.size
                    header_areas.append({
                        'top': loc['y'],
                        'bottom': loc['y'] + sz['height'],
                        'left': loc['x'],
                        'right': loc['x'] + sz['width']
                    })
                
                for elem in download_elements:
                    elem_loc = elem.location
                    # 检查是否在header区域
                    in_header = False
                    for header_area in header_areas:
                        if (header_area['top'] <= elem_loc['y'] <= header_area['bottom'] and
                            header_area['left'] <= elem_loc['x'] <= header_area['right']):
                            in_header = True
                            break
                    
                    if not in_header:
                        # 获取href或onclick中的URL
                        href = elem.get_attribute("href")
                        if not href:
                            onclick = elem.get_attribute("onclick")
                            # 从onclick中提取URL（如果需要）
                            if onclick and "download" in onclick.lower():
                                # 尝试从onclick中提取URL
                                pass
                        
                        if href:
                            download_url = href
                            logger.info(f"找到下载链接（方法2）: {download_url}")
                            break
                        
            except TimeoutException:
                pass
            except Exception as e:
                logger.warning(f"查找中文'下载'按钮失败: {str(e)}")
        
        # 方法3: 查找文件列表区域中带有下载功能的链接（通过class="name"的链接）
        if not download_url:
            try:
                # 根据用户提供的HTML，下载链接在class="name"的元素中
                # 查找当前文件对应的下载链接
                name_links = driver.find_elements(By.CSS_SELECTOR, "a.name[href*='/download?path=']")
                for link in name_links:
                    href = link.get_attribute("href")
                    if href:
                        # 检查href是否包含文件名（处理空格编码）
                        filename_encoded_space = filename.replace(' ', '%20')
                        filename_encoded_plus = filename.replace(' ', '+')
                        if (filename in href or 
                            filename_encoded_space in href or 
                            filename_encoded_plus in href):
                            # 确保不在header中
                            link_loc = link.location
                            header_elements = driver.find_elements(By.CSS_SELECTOR, "header, .header, #header")
                            in_header = False
                            for header in header_elements:
                                header_loc = header.location
                                header_sz = header.size
                                if (header_loc['y'] <= link_loc['y'] <= header_loc['y'] + header_sz['height']):
                                    in_header = True
                                    break
                            
                            if not in_header:
                                download_url = href
                                logger.info(f"找到下载链接（方法3）: {download_url}")
                                break
                        # 额外检查：URL解码后是否包含文件名
                        try:
                            from urllib.parse import unquote
                            decoded_href = unquote(href)
                            if filename in decoded_href:
                                link_loc = link.location
                                header_elements = driver.find_elements(By.CSS_SELECTOR, "header, .header, #header")
                                in_header = False
                                for header in header_elements:
                                    header_loc = header.location
                                    header_sz = header.size
                                    if (header_loc['y'] <= link_loc['y'] <= header_loc['y'] + header_sz['height']):
                                        in_header = True
                                        break
                                
                                if not in_header:
                                    download_url = href
                                    logger.info(f"找到下载链接（方法3，URL解码匹配）: {download_url}")
                                    break
                        except:
                            pass
            except Exception as e:
                logger.warning(f"通过name链接查找下载按钮失败: {str(e)}")
        
        if not download_url:
            logger.error(f"无法找到文件列表中的下载链接: {filename}")
            return False
        
        # 使用JavaScript触发下载，交给Chrome下载器（不直接点击链接）
        try:
            # 方法：创建一个隐藏的下载链接并触发下载
            driver.execute_script(f"""
                var link = document.createElement('a');
                link.href = '{download_url}';
                link.download = '{filename}';
                link.style.display = 'none';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            """)
            logger.info(f"已触发Chrome下载器下载: {filename}")
            download_button_found = True
        except Exception as e:
            logger.warning(f"JavaScript触发下载失败，尝试直接导航: {str(e)}")
            # 备用方法：直接导航到下载URL（Chrome会自动下载）
            try:
                driver.get(download_url)
                download_button_found = True
                logger.info(f"已通过导航触发下载: {filename}")
                # 需要等待一下，然后返回上一页或刷新文件列表
                time.sleep(1)
            except Exception as e2:
                logger.error(f"导航到下载URL失败: {str(e2)}")
                return False
        
        # 等待一小段时间让下载开始
        time.sleep(2)
        
        # 取消选中文件（为下一个文件做准备）
        try:
            checkbox = file_element.find_element(By.CSS_SELECTOR, "input[type='checkbox'], .select-checkbox, .checkbox")
            if checkbox.is_selected():
                checkbox.click()
        except:
            pass  # 如果没有复选框，忽略
        
        return True
        
    except Exception as e:
        logger.error(f"下载文件 {filename} 时出错: {str(e)}")
        return False


def download_file(driver: webdriver.Chrome, file_info: Tuple[str, str, str]) -> bool:
    """
    导航到文件目录，选中文件，然后点击Download按钮触发下载（用于重新下载场景）
    """
    current_path, filename, element_xpath = file_info
    
    try:
        # 如果需要，导航到文件所在目录
        if current_path:
            logger.info(f"导航到目录: {current_path}")
            navigate_to_directory(driver, current_path)
            wait_for_page_load(driver)
            time.sleep(2)
        
        # 使用当前目录下载函数
        return download_file_in_current_directory(driver, file_info)
        
    except Exception as e:
        logger.error(f"下载文件 {filename} 时出错: {str(e)}")
        return False


def navigate_to_directory(driver: webdriver.Chrome, target_path: str) -> bool:
    """
    导航到指定目录（通过breadcrumb或URL）
    """
    try:
        # 获取当前breadcrumb路径
        breadcrumbs = driver.find_elements(By.CSS_SELECTOR, ".breadcrumb a, .crumb a, nav a")
        current_breadcrumb = []
        
        for bc in breadcrumbs:
            current_breadcrumb.append(bc.text.strip())
        
        # 解析目标路径
        target_parts = target_path.split("/") if "/" in target_path else target_path.split("\\")
        target_parts = [p for p in target_parts if p]
        
        # 如果已经在目标目录，直接返回
        if len(current_breadcrumb) >= len(target_parts):
            match = True
            for i, part in enumerate(target_parts):
                if i + 1 < len(current_breadcrumb) and current_breadcrumb[i + 1] != part:
                    match = False
                    break
            if match:
                return True
        
        # 否则，需要从根目录开始导航
        # 先返回根目录
        while len(breadcrumbs) > 1:
            navigate_back(driver)
            breadcrumbs = driver.find_elements(By.CSS_SELECTOR, ".breadcrumb a, .crumb a, nav a")
            time.sleep(1)
        
        # 然后逐级进入目标目录
        for part in target_parts:
            file_list = get_file_list(driver)
            for name, is_folder, element in file_list:
                if name == part and is_folder:
                    clickable = element.find_element(By.CSS_SELECTOR, "a, .name, .filename")
                    clickable.click()
                    time.sleep(2)
                    wait_for_page_load(driver)
                    break
        
        return True
        
    except Exception as e:
        logger.error(f"导航到目录 {target_path} 失败: {str(e)}")
        return False


def generate_failure_report(failed_files: List[Tuple[str, str, str]]) -> None:
    """
    生成下载失败报告
    """
    try:
        report_path = os.path.join(os.path.dirname(__file__), "download_failures.txt")
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("ownCloud 下载失败报告\n")
            f.write("=" * 80 + "\n")
            f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"失败文件总数: {len(failed_files)}\n")
            f.write("=" * 80 + "\n\n")
            
            for i, file_info in enumerate(failed_files, 1):
                current_path, filename, _ = file_info
                file_path = os.path.join(current_path, filename).replace("\\", "/") if current_path else filename
                f.write(f"{i}. {file_path}\n")
            
            f.write("\n" + "=" * 80 + "\n")
            f.write("报告结束\n")
            f.write("=" * 80 + "\n")
        
        logger.info(f"下载失败报告已生成: {report_path}")
        
    except Exception as e:
        logger.error(f"生成下载失败报告时出错: {str(e)}")


def retry_download(driver: webdriver.Chrome, file_info: Tuple[str, str, str], max_retries: int = None) -> bool:
    """
    下载文件，包含重试逻辑（最多10次）
    """
    if max_retries is None:
        max_retries = config.MAX_RETRIES
    
    current_path, filename, element_xpath = file_info
    
    for attempt in range(max_retries):
        try:
            logger.info(f"尝试下载 {filename} (第 {attempt + 1}/{max_retries} 次)")
            
            # 触发下载
            if download_file(driver, file_info):
                # 监控下载并传递目标目录
                if monitor_download(driver, filename, current_path):
                    logger.info(f"成功下载: {filename}")
                    return True
                else:
                    logger.warning(f"下载监控失败: {filename}")
            else:
                logger.warning(f"触发下载失败: {filename}")
            
            # 如果失败，等待后重试
            if attempt < max_retries - 1:
                wait_time = config.RETRY_WAIT_TIME * (attempt + 1)  # 递增等待时间
                logger.info(f"等待 {wait_time} 秒后重试...")
                time.sleep(wait_time)
                
        except Exception as e:
            logger.error(f"下载尝试 {attempt + 1} 出错: {str(e)}")
            if attempt < max_retries - 1:
                time.sleep(config.RETRY_WAIT_TIME)
    
    logger.error(f"下载失败，已重试 {max_retries} 次: {filename}")
    return False


def main():
    """
    主函数：按顺序执行初始化、登录、扫描、下载
    """
    driver = None
    
    try:
        logger.info("=" * 60)
        logger.info("ownCloud自动化下载工具启动")
        logger.info("=" * 60)
        
        # 1. 初始化ChromeDriver
        logger.info("步骤1: 初始化ChromeDriver...")
        driver = setup_chrome_driver()
        
        # 2. 登录ownCloud
        logger.info("步骤2: 登录ownCloud...")
        if not login_to_owncloud(driver):
            logger.error("登录失败，程序退出")
            return
        
        # 3. 检查本地目录结构
        logger.info("步骤3: 检查本地目录结构...")
        download_base = os.path.abspath(config.DOWNLOAD_DIR)
        if os.path.exists(download_base):
            logger.info(f"本地下载目录已存在: {download_base}")
        else:
            logger.info(f"创建本地下载目录: {download_base}")
            os.makedirs(download_base, exist_ok=True)
        
        # 4. 循环执行扫描和下载，直到所有文件都下载成功
        global failed_files, download_stats
        cycle_count = 0
        max_cycles = config.MAX_FULL_CYCLES
        
        while cycle_count < max_cycles:
            cycle_count += 1
            logger.info("=" * 60)
            logger.info(f"第 {cycle_count} 轮完整下载流程")
            logger.info("=" * 60)
            
            # 重置失败列表和统计信息
            failed_files = []
            download_stats = {
                'total_files': 0,
                'existing_files': 0,
                'downloaded': 0,
                'failed': 0
            }
            
            # 扫描目录结构并立即下载文件（自动跳过已存在的文件）
            logger.info(f"步骤4.{cycle_count}: 扫描目录结构并下载文件（自动跳过已存在的文件）...")
            scan_directory(driver)
            
            logger.info(f"本轮扫描和下载完成，成功下载的文件已保存，失败的文件数: {len(failed_files)}")
            
            # 如果有失败的文件，先尝试重试下载
            if failed_files:
                logger.info("=" * 60)
                logger.info(f"步骤5.{cycle_count}: 开始重试下载失败的文件（共 {len(failed_files)} 个）...")
                logger.info("=" * 60)
                
                retry_failed_files = []
                for i, file_info in enumerate(failed_files, 1):
                    current_path, filename, _ = file_info
                    file_path = os.path.join(current_path, filename).replace("\\", "/") if current_path else filename
                    
                    logger.info(f"[重试 {i}/{len(failed_files)}] {file_path}")
                    
                    # 检查文件是否在重试期间已经下载成功（可能其他进程或手动下载）
                    local_file_path = os.path.join(config.DOWNLOAD_DIR, current_path, filename) if current_path else os.path.join(config.DOWNLOAD_DIR, filename)
                    local_file_path = os.path.normpath(local_file_path)
                    if os.path.exists(local_file_path) and os.path.isfile(local_file_path):
                        logger.info(f"文件已存在，跳过重试: {file_path}")
                        continue
                    
                    # 重试下载（使用download_file函数，因为需要导航到目录）
                    if retry_download(driver, file_info):
                        logger.info(f"重试下载成功: {file_path}")
                    else:
                        logger.error(f"重试下载仍然失败: {file_path}")
                        retry_failed_files.append(file_info)
                
                failed_files = retry_failed_files  # 更新失败列表
            
            # 输出本轮统计信息
            logger.info("=" * 60)
            logger.info(f"第 {cycle_count} 轮统计信息:")
            logger.info(f"  - 发现文件总数: {download_stats['total_files']}")
            logger.info(f"  - 已存在文件数: {download_stats['existing_files']}")
            logger.info(f"  - 本轮下载成功: {download_stats['downloaded']}")
            logger.info(f"  - 本轮下载失败: {download_stats['failed']}")
            logger.info(f"  - 当前失败列表: {len(failed_files)} 个文件")
            logger.info("=" * 60)
            
            # 判断是否继续循环
            # 条件1：如果没有失败的文件，且本轮没有新下载，说明所有文件都已完成
            if not failed_files and download_stats['downloaded'] == 0:
                logger.info("=" * 60)
                logger.info(f"所有文件下载成功！共执行了 {cycle_count} 轮下载流程")
                logger.info("=" * 60)
                break
            
            # 条件2：如果还有失败的文件，且未达到最大循环次数，等待后继续下一轮
            if cycle_count < max_cycles:
                if failed_files or download_stats['downloaded'] > 0:
                    logger.info("=" * 60)
                    logger.warning(f"仍有 {len(failed_files)} 个文件下载失败或有新文件需下载，将在 {config.CYCLE_WAIT_TIME} 秒后开始第 {cycle_count + 1} 轮下载...")
                    logger.info("=" * 60)
                    time.sleep(config.CYCLE_WAIT_TIME)
                    
                    # 重新登录（避免会话过期）
                    logger.info("重新登录以确保会话有效...")
                    if not login_to_owncloud(driver):
                        logger.error("重新登录失败，继续使用当前会话")
        
        # 6. 生成下载报告
        logger.info("=" * 60)
        logger.info("下载任务完成！")
        logger.info(f"总共执行了 {cycle_count} 轮完整下载流程")
        if failed_files:
            logger.error(f"最终下载失败的文件数: {len(failed_files)}")
            generate_failure_report(failed_files)
        else:
            logger.info("所有文件下载成功！")
        logger.info("=" * 60)
        
    except Exception as e:
        logger.error(f"程序执行过程中发生错误: {str(e)}", exc_info=True)
        
    finally:
        if driver:
            logger.info("关闭浏览器...")
            driver.quit()


if __name__ == "__main__":
    main()