"""
统一下载工具入口
支持 ownCloud 和 SharePoint
"""

import argparse
import sys
import config

def main():
    parser = argparse.ArgumentParser(description='云存储文件下载工具')
    parser.add_argument('--source', '-s', 
                        choices=['owncloud', 'sharepoint', 'auto'],
                        default='auto',
                        help='选择下载源类型 (owncloud/sharepoint)')
    
    args = parser.parse_args()
    
    source = args.source
    
    if source == 'auto':
        # 自动检测逻辑
        # 优先检查 SHAREPOINT_URL
        sp_url = getattr(config, 'SHAREPOINT_URL', '')
        oc_url = getattr(config, 'OWNCLOUD_URL', '')
        
        if sp_url and sp_url != "你的SharePoint分享链接":
            source = 'sharepoint'
        elif oc_url and oc_url != "你得到的分享链接":
            if 'sharepoint.com' in oc_url or 'onedrive' in oc_url:
                source = 'sharepoint'
            else:
                source = 'owncloud'
        else:
            print("错误: 未在 config.py 中检测到有效的分享链接。")
            sys.exit(1)
            
    print(f"正在检测并启动 {source} 下载模块...")
    
    try:
        if source == 'owncloud':
            from main import main as owncloud_main
            owncloud_main()
        elif source == 'sharepoint':
            from sharepoint_download import main as sharepoint_main
            sharepoint_main()
    except KeyboardInterrupt:
        print("\n程序已被用户中断。")
    except Exception as e:
        print(f"\n运行过程中发生错误: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
