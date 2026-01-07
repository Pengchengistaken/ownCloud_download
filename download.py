"""
统一下载工具入口 - 交互模式
支持 ownCloud 和 SharePoint
"""

import sys
import os
import re

def detect_source_type(url: str) -> str:
    """根据URL自动检测来源类型"""
    url_lower = url.lower()
    if 'sharepoint.com' in url_lower or 'onedrive' in url_lower:
        return 'sharepoint'
    else:
        return 'owncloud'

def main():
    print("=" * 60)
    print("云存储自动化下载工具")
    print("支持: ownCloud | SharePoint")
    print("=" * 60)
    print()
    
    # 交互式输入分享链接
    url = input("请输入分享链接: ").strip()
    
    if not url:
        print("错误: 分享链接不能为空")
        sys.exit(1)
    
    # 自动检测来源类型
    source = detect_source_type(url)
    print(f"\n检测到来源类型: {source.upper()}")
    
    # 询问下载目录
    download_dir = input(f"\n请输入下载目录 (默认: ./downloads): ").strip()
    if not download_dir:
        download_dir = "./downloads"
    
    # 创建下载目录
    os.makedirs(download_dir, exist_ok=True)
    print(f"下载目录: {os.path.abspath(download_dir)}")
    
    # 根据来源类型获取额外信息
    password = ""
    if source == 'owncloud':
        password = input("\n请输入分享密码 (如无密码直接回车): ").strip()
    
    print("\n" + "=" * 60)
    print("配置信息:")
    print(f"  来源: {source.upper()}")
    print(f"  链接: {url}")
    if password:
        print(f"  密码: {'*' * len(password)}")
    print(f"  下载目录: {os.path.abspath(download_dir)}")
    print("=" * 60)
    
    # 确认开始
    confirm = input("\n是否开始下载? (Y/n): ").strip().lower()
    if confirm and confirm not in ['y', 'yes', '是']:
        print("已取消下载")
        sys.exit(0)
    
    print("\n开始下载...")
    
    try:
        if source == 'owncloud':
            from main import main as owncloud_main
            # 临时设置配置
            import config
            config.OWNCLOUD_URL = url
            config.SHARE_PASSWORD = password
            config.DOWNLOAD_DIR = download_dir
            owncloud_main()
        elif source == 'sharepoint':
            from sharepoint_download import SharePointDownloader, download_stats
            import config
            config.SHAREPOINT_URL = url
            config.SHAREPOINT_DOWNLOAD_DIR = download_dir
            
            downloader = SharePointDownloader()
            try:
                if downloader.access_share_link(url):
                    downloader.traverse_and_download()
                    
                    # 尝试重试失败的文件
                    failed_count = len(downloader.download_state.get_failed_files())
                    if failed_count > 0:
                        print(f"\n首轮下载完成，{failed_count} 个文件失败，开始重试...")
                        retried = downloader.retry_failed_downloads(url)
                        print(f"重试完成，成功 {retried} 个文件")
                    
                    print("\n下载流程结束")
                    print(f"统计信息: {download_stats}")
                    
                    # 显示仍然失败的文件
                    remaining_failed = downloader.download_state.get_failed_files()
                    if remaining_failed:
                        print(f"\n警告: 仍有 {len(remaining_failed)} 个文件下载失败")
                        for f in remaining_failed[:5]:
                            print(f"  - {os.path.join(f['path'], f['name'])}")
                        if len(remaining_failed) > 5:
                            print(f"  ... 还有 {len(remaining_failed) - 5} 个文件")
                        print("提示: 再次运行程序可重试失败的下载")
                    else:
                        downloader.download_state.clear_state()
                        print("\n所有文件下载成功！")
            finally:
                downloader.close()
                
    except KeyboardInterrupt:
        print("\n\n程序已被用户中断。")
    except Exception as e:
        print(f"\n运行过程中发生错误: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
