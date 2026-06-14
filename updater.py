# -*- coding: utf-8 -*-
"""
自动更新模块 - 检查GitHub Releases并提示用户更新
"""
import json
import urllib.request
import webbrowser
from packaging import version

# 应用当前版本
CURRENT_VERSION = "1.0.0"

# GitHub仓库信息（发布后需要更新）
GITHUB_REPO = "your-username/file-converter"  # 格式: owner/repo
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
RELEASE_PAGE_URL = f"https://github.com/{GITHUB_REPO}/releases/latest"


def check_for_updates(timeout=5):
    """
    检查GitHub Releases是否有新版本

    Returns:
        tuple: (has_update: bool, latest_version: str, download_url: str, release_notes: str)
    """
    try:
        req = urllib.request.Request(
            GITHUB_API_URL,
            headers={'User-Agent': 'FileConverter-UpdateChecker'}
        )

        with urllib.request.urlopen(req, timeout=timeout) as response:
            data = json.loads(response.read().decode('utf-8'))

            latest_version = data.get('tag_name', '').lstrip('v')
            release_notes = data.get('body', '暂无更新说明')

            # 查找Windows可执行文件下载链接
            download_url = RELEASE_PAGE_URL
            for asset in data.get('assets', []):
                if asset['name'].endswith('.exe'):
                    download_url = asset['browser_download_url']
                    break

            # 比较版本号
            if version.parse(latest_version) > version.parse(CURRENT_VERSION):
                return True, latest_version, download_url, release_notes
            else:
                return False, CURRENT_VERSION, '', ''

    except Exception as e:
        # 网络错误或API限制，静默失败
        print(f"更新检查失败: {e}")
        return False, CURRENT_VERSION, '', ''


def open_download_page(url):
    """在浏览器中打开下载页面"""
    webbrowser.open(url)


def get_current_version():
    """获取当前应用版本"""
    return CURRENT_VERSION
