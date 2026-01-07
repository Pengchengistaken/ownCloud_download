"""
ownCloud自动化下载配置文件
"""

# ownCloud分享链接
OWNCLOUD_URL = "你得到的分享链接"

# 分享密码
SHARE_PASSWORD = "你得到的分享密码"

# 本地下载目录（相对于项目根目录）
DOWNLOAD_DIR = "./downloads"

# 最大重试次数
MAX_RETRIES = 10

# 单文件下载超时时间（秒）
DOWNLOAD_TIMEOUT = 3600  # 1小时

# 页面加载超时时间（秒）
PAGE_LOAD_TIMEOUT = 60

# 下载重试等待时间（秒）
RETRY_WAIT_TIME = 30

# 最大完整循环次数（整个扫描和下载流程的重复次数）
MAX_FULL_CYCLES = 10

# 每轮循环之间的等待时间（秒）
CYCLE_WAIT_TIME = 60


# =======================
# SharePoint 配置
# =======================

# SharePoint 分享链接
SHAREPOINT_URL = "你的SharePoint分享链接"

# 本地下载目录（可以和ownCloud共用，也可以单独设置）
SHAREPOINT_DOWNLOAD_DIR = "./downloads"

# SharePoint 特定配置
SHAREPOINT_PAGE_LOAD_TIMEOUT = 60  # 页面加载超时（秒）
SHAREPOINT_DOWNLOAD_TIMEOUT = 3600  # 单文件下载超时（秒）
SHAREPOINT_MAX_RETRIES = 10  # 单文件最大重试次数
SHAREPOINT_RETRY_WAIT_TIME = 30  # 重试等待时间（秒）
SHAREPOINT_MAX_FULL_CYCLES = 10  # 最大完整循环次数
SHAREPOINT_CYCLE_WAIT_TIME = 60  # 循环间等待时间（秒）
