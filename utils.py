#!/usr/bin/env python
# -*- coding: utf-8 -*-
import logging
import importlib.util
from pathlib import Path

# 設定日誌
def setup_logging():
    from pathlib import Path
    log_dir = Path.home() / '.libreoffice' / 'speech_to_text_logs'
    try:
        # 確保日誌目錄存在
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / 'speech_extension.log'
        
        # 使用 logging.getLogger() 獲取日誌對象
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)
        
        # 使用 FileHandler 確定編碼方式為 UTF-8
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        # 設定日誌格式
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        
        # 清除舊的 handler 并添加新的 handler
        logger.handlers = []
        logger.addHandler(file_handler)
        
        return log_dir
    except Exception as e:
        # 如果無法創建日誌檔案，則使用標準輸出
        logging.basicConfig(level=logging.DEBUG)
        logging.warning(f"Could not create log file: {e}")
        return None

# 檢查模組是否已安裝，支持檢查用戶虛擬環境中的模組
def check_module_installed(module_name):
    try:
        # 首先嘗試直接導入檢查
        spec = importlib.util.find_spec(module_name)
        if spec is not None:
            logging.debug(f"檢查模組 {module_name}: 已安裝")
            return True
            
        # 如果直接導入失敗，檢查用戶虛擬環境中是否存在
        import platform
        from pathlib import Path
        home_path = Path.home()
        
        # 根據操作系統選擇正確的路徑格式
        system = platform.system()
        if system == "Windows":
            # Windows 路徑
            venv_site_packages = home_path / '.libreoffice' / 'python_env' / 'venv' / 'Lib' / 'site-packages'
        else:
            # Linux/Mac 路徑 - 需要檢測 Python 版本
            import sys
            python_major = sys.version_info.major
            python_minor = sys.version_info.minor
            venv_site_packages = home_path / '.libreoffice' / 'python_env' / 'venv' / 'lib' / f'python{python_major}.{python_minor}' / 'site-packages'
            
            # 如果這個路徑不存在，嘗試一般的路徑
            if not venv_site_packages.exists():
                # 嘗試各種可能的 Python 版本目錄
                for minor_ver in range(python_minor - 1, python_minor + 2):
                    if minor_ver > 0:  # 確保次要版本號有效
                        alt_path = home_path / '.libreoffice' / 'python_env' / 'venv' / 'lib' / f'python{python_major}.{minor_ver}' / 'site-packages'
                        if alt_path.exists():
                            venv_site_packages = alt_path
                            break
        
        logging.debug(f"檢查虛擬環境路徑: {venv_site_packages}")
        
        # 檢查模組是否存在於虛擬環境的目錄中
        if venv_site_packages.exists():
            # 檢查模組目錄
            module_dir = venv_site_packages / module_name
            if module_dir.exists() and module_dir.is_dir():
                logging.debug(f"檢查模組 {module_name}: 存在於用戶虛擬環境中")
                return True
                
            # 檢查.py文件
            module_file = venv_site_packages / f"{module_name}.py"
            if module_file.exists():
                logging.debug(f"檢查模組 {module_name}: 存在於用戶虛擬環境中")
                return True
                
            # 檢查egg-info或dist-info目錄
            for info_dir in venv_site_packages.glob(f"{module_name}*egg-info") or venv_site_packages.glob(f"{module_name}*dist-info"):
                if info_dir.exists():
                    logging.debug(f"檢查模組 {module_name}: 存在於用戶虛擬環境中")
                    return True
                    
        logging.debug(f"檢查模組 {module_name}: 未安裝")
        return False
    except ImportError:
        logging.debug(f"檢查模組 {module_name} 時發生錯誤")
        return False

# 顯示訊息對話框
def show_message_box(ctx, message, title, msg_type, buttons=1):
    sm = ctx.getServiceManager()
    toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent = toolkit.getDesktopWindow()
    mb = toolkit.createMessageBox(parent, msg_type, buttons, title, message)
    return mb.execute()