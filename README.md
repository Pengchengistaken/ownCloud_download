# ownCloud 自动化下载工具

一个基于 Selenium 的自动化工具，用于从 ownCloud 共享文件夹中批量下载文件。支持递归遍历所有文件夹，自动创建本地目录结构，并具备断点续传、自动重试等功能。

## 功能特性

- ✅ **自动登录**：使用分享链接和密码自动登录 ownCloud
- ✅ **递归遍历**：自动遍历所有文件夹层级，创建对应的本地目录结构
- ✅ **智能下载**：逐个下载文件，等待每个文件完成后再继续
- ✅ **断点续传**：自动检测已存在的文件，跳过已下载的内容
- ✅ **自动重试**：单个文件下载失败时自动重试（最多10次）
- ✅ **循环执行**：自动循环执行完整下载流程，直到所有文件下载成功
- ✅ **失败报告**：生成下载失败文件的详细报告
- ✅ **动态加载**：自动滚动页面，确保加载所有文件
- ✅ **会话保持**：自动重新登录，避免会话过期

## 系统要求

- Python 3.7+
- Chrome 浏览器（最新版本）
- 稳定的网络连接

## 安装步骤

### 1. 克隆或下载项目

```bash
git clone <repository-url>
cd ownCloud_download
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 配置参数

编辑 `config.py` 文件，设置以下参数：

```python
# ownCloud分享链接
OWNCLOUD_URL = "你得到的分享链接"

# 分享密码
SHARE_PASSWORD = "你得到的分享密码"

# 本地下载目录（相对于项目根目录）
DOWNLOAD_DIR = "./downloads"
```

### 4. 运行程序

```bash
python main.py
```

## 配置说明

在 `config.py` 中可以调整以下参数：

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `OWNCLOUD_URL` | ownCloud 分享链接 | 必填 |
| `SHARE_PASSWORD` | 分享密码 | 必填 |
| `DOWNLOAD_DIR` | 本地下载目录 | `"./downloads"` |
| `MAX_RETRIES` | 单个文件最大重试次数 | `10` |
| `DOWNLOAD_TIMEOUT` | 单文件下载超时时间（秒） | `3600` (1小时) |
| `PAGE_LOAD_TIMEOUT` | 页面加载超时时间（秒） | `60` |
| `RETRY_WAIT_TIME` | 下载重试等待时间（秒） | `30` |
| `MAX_FULL_CYCLES` | 最大完整循环次数 | `10` |
| `CYCLE_WAIT_TIME` | 每轮循环之间的等待时间（秒） | `60` |

## 使用方法

### 基本使用

1. 确保已正确配置 `config.py` 中的 `OWNCLOUD_URL` 和 `SHARE_PASSWORD`
2. 运行程序：
   ```bash
   python main.py
   ```
3. 程序会自动：
   - 打开 Chrome 浏览器
   - 登录 ownCloud
   - 扫描所有文件夹
   - 下载缺失的文件
   - 自动重试失败的文件
   - 循环执行直到所有文件下载完成

### 工作流程

程序执行流程如下：

1. **初始化**：启动 ChromeDriver，配置下载路径
2. **登录**：使用分享链接和密码登录 ownCloud
3. **扫描下载**：
   - 递归遍历所有文件夹
   - 检查文件是否已存在本地
   - 下载缺失的文件
   - 自动重试失败的文件
4. **循环执行**：
   - 如果还有失败的文件，等待后重新扫描
   - 最多执行 10 轮完整流程
   - 每轮之间自动重新登录保持会话
5. **生成报告**：生成 `download_failures.txt` 记录最终失败的文件

### 日志文件

程序运行时会生成以下文件：

- `owncloud_download.log`：详细的运行日志
- `download_failures.txt`：最终下载失败的文件列表（如果有）

## 注意事项

1. **网络稳定性**：由于需要从德国服务器下载，建议在网络稳定时运行
2. **磁盘空间**：确保有足够的磁盘空间（本项目需要 17GB+）
3. **浏览器窗口**：程序运行时会打开 Chrome 浏览器窗口，请勿手动操作
4. **中断恢复**：程序支持断点续传，可以随时中断，重新运行时会自动跳过已下载的文件
5. **下载路径**：下载的文件会保存在 `config.DOWNLOAD_DIR` 指定的目录中，保持原有的文件夹结构

## 故障排除

### 问题：无法找到 ChromeDriver

**解决方案**：程序使用 `webdriver-manager` 自动管理 ChromeDriver，首次运行时会自动下载。如果遇到问题，请确保：
- Chrome 浏览器已正确安装
- 网络连接正常（需要下载 ChromeDriver）

### 问题：下载超时

**解决方案**：
- 检查网络连接
- 增加 `DOWNLOAD_TIMEOUT` 的值（在 `config.py` 中）
- 程序会自动重试，耐心等待

### 问题：页面加载失败

**解决方案**：
- 检查 `OWNCLOUD_URL` 是否正确
- 检查网络连接
- 增加 `PAGE_LOAD_TIMEOUT` 的值

### 问题：文件下载失败

**解决方案**：
- 查看 `owncloud_download.log` 了解详细错误信息
- 程序会自动重试失败的文件
- 最终失败的文件会记录在 `download_failures.txt` 中

## 项目结构

```
ownCloud_download/
├── main.py              # 主程序文件
├── config.py            # 配置文件
├── requirements.txt     # Python 依赖
├── README.md           # 项目说明文档
├── .gitignore          # Git 忽略文件
├── downloads/          # 下载目录（自动创建）
├── owncloud_download.log    # 运行日志（自动生成）
└── download_failures.txt   # 失败报告（如果有）
```

## 技术栈

- **Selenium**：浏览器自动化框架
- **ChromeDriver**：Chrome 浏览器驱动
- **webdriver-manager**：自动管理 WebDriver 二进制文件

## 许可证

本项目仅供学习和个人使用。

## 贡献

欢迎提交 Issue 和 Pull Request！
