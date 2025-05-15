#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import os
import subprocess
import logging
from pathlib import Path
import platform

# 新增 - 修復 Linux 權限問題的函數
def fix_venv_permissions():
    """修復虛擬環境的權限問題，主要針對 Linux 系統"""
    try:
        # 檢測操作系統
        system = platform.system()
        if system != "Linux":
            return True  # 非 Linux 系統不需要修復權限
            
        logging.info("檢測並修復 Linux 虛擬環境權限")
        
        # 獲取用戶目錄
        home_path = Path.home()
        venv_dir = home_path / '.libreoffice' / 'python_env' / 'venv'
        
        if not venv_dir.exists():
            logging.info("虛擬環境不存在，無需修復權限")
            return True
            
        # 修復 bin 目錄下所有檔案的執行權限
        bin_dir = venv_dir / 'bin'
        if bin_dir.exists():
            logging.info(f"修復 {bin_dir} 下檔案的執行權限")
            for script in bin_dir.glob('*'):
                try:
                    os.chmod(str(script), 0o755)
                    logging.debug(f"已設置執行權限: {script}")
                except Exception as e:
                    logging.error(f"設置 {script} 執行權限時發生錯誤: {str(e)}")
                    
        # 修復 lib 目錄的讀取權限
        for lib_dir in venv_dir.glob('lib/*/site-packages'):
            if lib_dir.exists():
                try:
                    os.chmod(str(lib_dir), 0o755)
                    logging.debug(f"已設置目錄權限: {lib_dir}")
                    # 遞迴設置所有子目錄和文件的權限
                    for root, dirs, files in os.walk(str(lib_dir)):
                        for d in dirs:
                            try:
                                full_path = os.path.join(root, d)
                                os.chmod(full_path, 0o755)
                            except Exception as e:
                                logging.error(f"設置目錄 {d} 權限時發生錯誤: {str(e)}")
                                
                        for f in files:
                            try:
                                full_path = os.path.join(root, f)
                                # .py 文件需要執行權限
                                if f.endswith('.py'):
                                    os.chmod(full_path, 0o755)
                                else:
                                    os.chmod(full_path, 0o644)
                            except Exception as e:
                                logging.error(f"設置文件 {f} 權限時發生錯誤: {str(e)}")
                except Exception as e:
                    logging.error(f"處理目錄 {lib_dir} 權限時發生錯誤: {str(e)}")
        
        logging.info("Linux 虛擬環境權限修復完成")
        return True
    except Exception as e:
        logging.error(f"修復權限時發生錯誤: {str(e)}")
        return False

# 導入工具函數
from utils import show_message_box, check_module_installed

# 直接在Python中執行安裝，避免版本不兼容
def install_modules_directly(ctx):
    """自動化安裝匹配版本的 Python 與必要模組，使用用戶目錄避免權限問題"""
    try:
        # 確保所有必要的模組都已導入
        import sys
        import os
        import subprocess
        import shutil
        import tempfile
        import re
        import urllib.request
        from pathlib import Path
        
        # 檢測操作系統
        system = platform.system()
        is_windows = system == "Windows"
        is_macos = system == "Darwin"
        is_linux = system == "Linux"
        
        # 如果是 Windows，則導入 winreg
        if is_windows:
            import winreg
            import zipfile
        
        # 創建日誌記錄
        logging.info(f"開始安裝必要模組，系統類型: {system}")
        
        # 確保必要目錄存在
        home_path = Path.home()
        flask_dir = home_path / '.libreoffice' / 'speech_api'
        log_dir = home_path / '.libreoffice' / 'speech_to_text_logs'
        temp_dir = home_path / '.libreoffice' / 'temp_installation'
        user_python_dir = home_path / '.libreoffice' / 'python_env'
        
        # 創建目錄 (無論安裝是否成功)
        flask_dir.mkdir(parents=True, exist_ok=True)
        log_dir.mkdir(parents=True, exist_ok=True)
        temp_dir.mkdir(parents=True, exist_ok=True)
        user_python_dir.mkdir(parents=True, exist_ok=True)
        
        # 顯示安裝開始訊息
        show_message_box(ctx, "正在檢測環境並準備安裝必要模組，請稍候...\n這可能需要幾分鐘時間。", "安裝準備中", 1)
        
        # 根據不同操作系統查找 LibreOffice 的 Python 版本
        lo_python_version = None
        
        if is_windows:
            # Windows 路徑
            lo_program_dir = Path("C:\\Program Files\\LibreOffice\\program")
            
            if lo_program_dir.exists():
                # 尋找 python-core-* 目錄
                python_cores = list(lo_program_dir.glob('python-core-*'))
                if python_cores:
                    for core_dir in python_cores:
                        # 從目錄名稱提取版本
                        match = re.search(r'python-core-(\d+\.\d+\.\d+)', str(core_dir))
                        if match:
                            lo_python_version = match.group(1)
                            logging.info(f"找到 LibreOffice Python 版本: {lo_python_version}")
                            break
                
                if not lo_python_version:
                    # 嘗試使用默認路徑提取版本
                    default_path = Path("C:\\Program Files\\LibreOffice\\program\\python-core-3.9.21")
                    if default_path.exists():
                        lo_python_version = "3.9.21"
                        logging.info(f"使用默認 LibreOffice Python 版本: {lo_python_version}")
        
        elif is_macos:
            # macOS 路徑
            lo_app_paths = [
                Path("/Applications/LibreOffice.app/Contents/Resources"),
                Path(home_path / "Applications/LibreOffice.app/Contents/Resources")
            ]
            
            for app_path in lo_app_paths:
                if app_path.exists():
                    # 查找 python-core-* 目錄
                    python_cores = list(app_path.glob('python-core-*'))
                    if python_cores:
                        for core_dir in python_cores:
                            match = re.search(r'python-core-(\d+\.\d+\.\d+)', str(core_dir))
                            if match:
                                lo_python_version = match.group(1)
                                logging.info(f"找到 LibreOffice Python 版本: {lo_python_version}")
                                break
            
            # 如果找不到特定版本，嘗試取得目前 Python 版本
            if not lo_python_version:
                try:
                    result = subprocess.run(
                        ["/usr/bin/python3", "-c", "import sys; print('.'.join(map(str, sys.version_info[:3])))"],
                        capture_output=True, 
                        text=True, 
                        check=True
                    )
                    lo_python_version = result.stdout.strip()
                    logging.info(f"使用系統 Python 版本: {lo_python_version}")
                except:
                    lo_python_version = "3.9.0"  # 默認版本
                    logging.info(f"無法檢測版本，使用默認版本: {lo_python_version}")
        
        elif is_linux:
            # Linux 路徑
            lo_paths = [
                Path("/usr/lib/libreoffice/program"),
                Path("/opt/libreoffice/program"),
                Path("/usr/lib64/libreoffice/program")  # 某些發行版
            ]
            
            for lo_path in lo_paths:
                if lo_path.exists():
                    # 查找 python-core-* 目錄
                    python_cores = list(lo_path.glob('python-core-*'))
                    if python_cores:
                        for core_dir in python_cores:
                            match = re.search(r'python-core-(\d+\.\d+\.\d+)', str(core_dir))
                            if match:
                                lo_python_version = match.group(1)
                                logging.info(f"找到 LibreOffice Python 版本: {lo_python_version}")
                                break
            
            # 如果找不到特定版本，嘗試取得目前 Python 版本
            if not lo_python_version:
                try:
                    result = subprocess.run(
                        ["python3", "-c", "import sys; print('.'.join(map(str, sys.version_info[:3])))"],
                        capture_output=True, 
                        text=True, 
                        check=True
                    )
                    lo_python_version = result.stdout.strip()
                    logging.info(f"使用系統 Python 版本: {lo_python_version}")
                except:
                    lo_python_version = "3.9.0"  # 默認版本
                    logging.info(f"無法檢測版本，使用默認版本: {lo_python_version}")
        
        # 如果無法檢測到版本，提供錯誤訊息
        if not lo_python_version:
            message = "無法找到 LibreOffice Python 版本。請確認 LibreOffice 已正確安裝。"
            logging.error("找不到 LibreOffice Python 版本")
            show_message_box(ctx, message, "安裝錯誤", 3)
            return False
        
        # 提取主要版本號（如 3.9)
        major_version = '.'.join(lo_python_version.split('.')[:2])
        logging.info(f"主要版本號: {major_version}")
        
        # 檢查和安裝匹配版本的 Python
        python_installed = False
        python_exe = None
        
        # 根據不同操作系統檢測 Python
        if is_windows:
            # 1. 檢查 Windows 註冊表中是否有匹配版本的 Python
            try:
                key_path = f"Software\\Python\\PythonCore\\{major_version}\\InstallPath"
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
                install_path = winreg.QueryValue(key, "")
                python_exe = Path(install_path) / "python.exe"
                if python_exe.exists():
                    logging.info(f"在註冊表找到匹配的 Python {major_version}: {python_exe}")
                    python_installed = True
            except:
                try:
                    key_path = f"Software\\Python\\PythonCore\\{major_version}\\InstallPath"
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    install_path = winreg.QueryValue(key, "")
                    python_exe = Path(install_path) / "python.exe"
                    if python_exe.exists():
                        logging.info(f"在註冊表找到匹配的 Python {major_version}: {python_exe}")
                        python_installed = True
                except:
                    logging.info(f"在註冊表沒有找到 Python {major_version}")
            
            # 2. 在常見 Windows 路徑尋找 Python
            if not python_installed:
                common_paths = [
                    Path(f"C:\\Python{major_version.replace('.', '')}\\python.exe"),
                    Path(f"C:\\Program Files\\Python{major_version.replace('.', '')}\\python.exe"),
                    Path(f"C:\\Program Files (x86)\\Python{major_version.replace('.', '')}\\python.exe"),
                    home_path / "AppData" / "Local" / "Programs" / "Python" / f"Python{major_version.replace('.', '')}" / "python.exe"
                ]
                
                for path in common_paths:
                    if path.exists():
                        python_exe = path
                        python_installed = True
                        logging.info(f"在常見路徑找到 Python {major_version}: {python_exe}")
                        break
        
        elif is_macos:
            # 在 macOS 上常見的 Python 路徑
            common_paths = [
                Path(f"/usr/local/bin/python{major_version}"),
                Path(f"/usr/bin/python{major_version}"),
                Path(f"/opt/homebrew/bin/python{major_version}"),
                Path(f"{home_path}/Library/Python/{major_version}/bin/python{major_version}")
            ]
            
            # 也檢查沒有版本號的 python3
            common_paths.append(Path("/usr/bin/python3"))
            common_paths.append(Path("/usr/local/bin/python3"))
            common_paths.append(Path("/opt/homebrew/bin/python3"))
            
            for path in common_paths:
                if path.exists():
                    # 檢查版本是否匹配
                    try:
                        result = subprocess.run(
                            [str(path), "-c", "import sys; print('.'.join(map(str, sys.version_info[:2])))"],
                            capture_output=True, 
                            text=True, 
                            check=True
                        )
                        version = result.stdout.strip()
                        if version == major_version or version.startswith(f"{major_version}."):
                            python_exe = path
                            python_installed = True
                            logging.info(f"在 macOS 找到匹配的 Python {major_version}: {python_exe}")
                            break
                    except:
                        logging.info(f"檢查版本時出錯: {path}")
        
        elif is_linux:
            # 在 Linux 上常見的 Python 路徑
            common_paths = [
                Path(f"/usr/bin/python{major_version}"),
                Path(f"/usr/local/bin/python{major_version}"),
                Path(f"{home_path}/.local/bin/python{major_version}")
            ]
            
            # 也檢查沒有版本號的 python3
            common_paths.append(Path("/usr/bin/python3"))
            common_paths.append(Path("/usr/local/bin/python3"))
            
            for path in common_paths:
                if path.exists():
                    # 檢查版本是否匹配
                    try:
                        result = subprocess.run(
                            [str(path), "-c", "import sys; print('.'.join(map(str, sys.version_info[:2])))"],
                            capture_output=True, 
                            text=True, 
                            check=True
                        )
                        version = result.stdout.strip()
                        if version == major_version or version.startswith(f"{major_version}."):
                            python_exe = path
                            python_installed = True
                            logging.info(f"在 Linux 找到匹配的 Python {major_version}: {python_exe}")
                            break
                    except:
                        logging.info(f"檢查版本時出錯: {path}")
        
        # 如果未找到匹配版本的 Python，則提示安裝
        if not python_installed:
            # 在不同作業系統上處理 Python 安裝
            if is_windows:
                # Windows 安裝流程
                try:
                    message = f"找不到與 LibreOffice 相容的 Python {major_version}，需要下載安裝。\n是否繼續？"
                    result = show_message_box(ctx, message, "下載 Python", 4, 3)
                    
                    if result != 2:
                        show_message_box(ctx, "安裝取消。必須有匹配版本的 Python 才能繼續。", "安裝取消", 1)
                        return False
                    
                    # 下載對應版本的 Python
                    show_message_box(ctx, f"正在下載 Python {major_version}，請稍候...\n這可能需要幾分鐘時間。", "下載中", 1)
                    
                    # 根據大版本選擇下載URL
                    if major_version == "3.9":
                        download_url = "https://www.python.org/ftp/python/3.9.13/python-3.9.13-amd64.exe"
                    elif major_version == "3.8":
                        download_url = "https://www.python.org/ftp/python/3.8.10/python-3.8.10-amd64.exe"
                    elif major_version == "3.10":
                        download_url = "https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe"
                    else:
                        # 默認下載 3.9 版本
                        download_url = "https://www.python.org/ftp/python/3.9.13/python-3.9.13-amd64.exe"
                        
                    installer_path = temp_dir / f"python_{major_version}_installer.exe"
                    
                    # 下載安裝檔
                    urllib.request.urlretrieve(download_url, installer_path)
                    logging.info(f"Python {major_version} 下載完成: {installer_path}")
                    
                    # 執行安裝程序 - 靜默安裝，只給當前用戶，添加到 PATH
                    show_message_box(ctx, f"正在安裝 Python {major_version}，請稍候...\n這可能需要幾分鐘時間。", "安裝中", 1)
                    install_args = [
                        str(installer_path),
                        "/quiet", 
                        "InstallAllUsers=0", 
                        "PrependPath=1", 
                        "Include_test=0",
                        "InstallLauncherAllUsers=0"
                    ]
                    
                    proc = subprocess.run(
                        install_args,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True
                    )
                    
                    if proc.returncode != 0:
                        logging.error(f"Python 安裝失敗: {proc.stderr}")
                        show_message_box(ctx, f"Python {major_version} 安裝失敗。請嘗試手動安裝。", "安裝錯誤", 3)
                        return False
                    
                    logging.info(f"Python {major_version} 安裝成功")
                    
                    # 查找新安裝的 Python
                    if major_version == "3.9":
                        python_exe = home_path / "AppData" / "Local" / "Programs" / "Python" / "Python39" / "python.exe"
                    elif major_version == "3.8":
                        python_exe = home_path / "AppData" / "Local" / "Programs" / "Python" / "Python38" / "python.exe"
                    elif major_version == "3.10":
                        python_exe = home_path / "AppData" / "Local" / "Programs" / "Python" / "Python310" / "python.exe"
                    
                    if not python_exe.exists():
                        # 嘗試在註冊表查詢安裝路徑
                        try:
                            key_path = f"Software\\Python\\PythonCore\\{major_version}\\InstallPath"
                            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path)
                            install_path = winreg.QueryValue(key, "")
                            python_exe = Path(install_path) / "python.exe"
                        except Exception as e:
                            logging.error(f"無法找到新安裝的 Python: {str(e)}")
                    
                    if not python_exe.exists():
                        logging.error("無法找到新安裝的 Python")
                        show_message_box(ctx, f"已安裝 Python {major_version}，但找不到執行檔。請重新啟動電腦後再試。", "安裝提示", 2)
                        return False
                        
                    python_installed = True
                    logging.info(f"使用新安裝的 Python: {python_exe}")
                    
                except Exception as e:
                    logging.error(f"下載/安裝 Python 時發生錯誤: {str(e)}")
                    show_message_box(ctx, f"下載/安裝 Python 時發生錯誤: {str(e)}", "安裝錯誤", 3)
                    return False
            
            elif is_macos:
                # macOS 安裝流程
                message = f"找不到與 LibreOffice 相容的 Python {major_version}。\n請使用 Homebrew 安裝：\n\nbrew install python@{major_version}\n\n或從 python.org 下載安裝，然後重試。"
                show_message_box(ctx, message, "需要安裝 Python", 2)
                return False
                
            elif is_linux:
                # Linux 安裝流程
                message = f"找不到與 LibreOffice 相容的 Python {major_version}。\n請使用您的發行版套件管理器安裝，例如：\n\n"
                message += f"Debian/Ubuntu: sudo apt install python{major_version} python{major_version}-venv python3-pip\n"
                message += f"Fedora: sudo dnf install python{major_version}\n\n"
                message += "安裝後請確保它們有正確的執行權限，可能需要執行:\n"
                message += f"chmod +x ~/.libreoffice/python_env/venv/bin/*\n\n"
                message += "然後重試。"
                show_message_box(ctx, message, "需要安裝 Python", 2)
                return False
        
        # 到這裡表示已找到匹配的 Python 版本
        logging.info(f"使用 Python 執行檔: {python_exe}")
        
        # 創建一個臨時目錄存放所有文件
        manual_install_dir = home_path / '.libreoffice' / 'python_modules'
        manual_install_dir.mkdir(parents=True, exist_ok=True)
        
        # 建立 requirements.txt 文件
        requirements_file = manual_install_dir / 'requirements.txt'
        with open(requirements_file, 'w') as f:
            f.write("SpeechRecognition\npyaudio\nflask\nrequests\n")
        
        # 使用匹配版本的 Python 安裝必要模組到臨時目錄
        show_message_box(ctx, f"正在使用 Python {major_version} 安裝必要模組，請稍候...", "模組安裝中", 1)
        
        # 創建虛擬環境 (使用用戶目錄避開權限問題)
        venv_dir = user_python_dir / "venv"
        if venv_dir.exists():
            shutil.rmtree(venv_dir)
        
        # 運行 Python 創建虛擬環境
        try:
            logging.info(f"創建虛擬環境: {venv_dir}")
            subprocess.run(
                [str(python_exe), "-m", "venv", str(venv_dir)],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            
            # 根據系統確定虛擬環境中的 Python 和 pip 路徑
            if is_windows:
                venv_python = venv_dir / "Scripts" / "python.exe"
                venv_pip = venv_dir / "Scripts" / "pip.exe"
            else:  # Linux/macOS
                venv_python = venv_dir / "bin" / "python"
                venv_pip = venv_dir / "bin" / "pip"
            
            if not venv_python.exists():
                logging.error(f"創建虛擬環境失敗，找不到 Python: {venv_python}")
                show_message_box(ctx, "創建 Python 虛擬環境失敗", "安裝錯誤", 3)
                return False
            
            # 在創建虛擬環境後，明確設置執行權限
            if not is_windows:  # Linux/Mac
                # 設置 Python 執行檔和其他腳本的執行權限
                try:
                    os.chmod(str(venv_python), 0o755)
                    # 還需要設置 pip 和其他腳本的執行權限
                    venv_bin_dir = venv_dir / "bin"
                    if venv_bin_dir.exists():
                        for script in venv_bin_dir.glob("*"):
                            os.chmod(str(script), 0o755)
                    logging.info(f"已設置虛擬環境執行權限: {venv_dir}")
                except Exception as e:
                    logging.error(f"設置執行權限時發生錯誤: {str(e)}")
                    show_message_box(ctx, f"無法設置執行權限，請嘗試手動運行:\nchmod +x {venv_python}", "權限錯誤", 3)
                    # 繼續執行，讓用戶有機會手動修復
            
            # 升級 pip
            logging.info("升級 pip")
            subprocess.run(
                [str(venv_python), "-m", "pip", "install", "--upgrade", "pip"],
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            
            # 安裝必要模組
            logging.info(f"安裝必要模組: {requirements_file}")
            result = subprocess.run(
                [str(venv_python), "-m", "pip", "install", "-r", str(requirements_file)],
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            if result.returncode != 0:
                logging.error(f"安裝模組失敗: {result.stderr}")
                # 在 Linux/macOS 上，可能需要安裝額外的依賴來獲取 PyAudio
                if is_linux and "pyaudio" in result.stderr.lower():
                    extra_message = "\n\n在 Linux 上，您可能需要先安裝 PortAudio 開發庫：\nsudo apt-get install portaudio19-dev python3-dev (Debian/Ubuntu)\nsudo dnf install portaudio-devel (Fedora)"
                    show_message_box(ctx, f"安裝必要模組失敗: \n{result.stderr}{extra_message}", "安裝錯誤", 3)
                elif is_macos and "pyaudio" in result.stderr.lower():
                    extra_message = "\n\n在 macOS 上，您可能需要先安裝 PortAudio：\nbrew install portaudio"
                    show_message_box(ctx, f"安裝必要模組失敗: \n{result.stderr}{extra_message}", "安裝錯誤", 3)
                else:
                    show_message_box(ctx, f"安裝必要模組失敗: \n{result.stderr}", "安裝錯誤", 3)
                return False
            
            logging.info("模組安裝成功")
            
            # 在 Linux 上再次確保所有檔案都有正確的權限
            if is_linux:
                try:
                    # 直接呼叫我們定義的權限修復函數
                    fix_venv_permissions()
                    logging.info("已再次檢查與修復虛擬環境權限")
                except Exception as e:
                    logging.error(f"修復虛擬環境權限時發生錯誤: {str(e)}")
                    # 繼續執行，因為我們已經安裝了模組
            
            # 找到虛擬環境中的 site-packages 目錄
            if is_windows:
                site_packages = list(venv_dir.glob('**/site-packages'))
            else:  # Linux/macOS
                site_packages = list(venv_dir.glob('lib/*/site-packages'))
                
            if not site_packages:
                logging.error("找不到虛擬環境中的 site-packages 目錄")
                show_message_box(ctx, "找不到安裝的模組", "安裝錯誤", 3)
                return False
                
            venv_site_packages = site_packages[0]
            logging.info(f"找到虛擬環境 site-packages: {venv_site_packages}")
            
            # 導入 api_service 模組來創建 API 腳本
            from api_service import create_api_script
            
            # 修改 speech_api.py 指向我們的虛擬環境
            create_api_script(flask_dir, venv_python)
            
            # 創建安裝完成標記
            config_marker = home_path / '.libreoffice' / 'speech_to_text_installed'
            config_marker.parent.mkdir(parents=True, exist_ok=True)
            with open(config_marker, 'w', encoding='utf-8') as f:
                f.write("Installed")
            
            show_message_box(ctx, f"已成功安裝必要模組到用戶級別的 Python 環境。\n\nAPI 服務已配置使用該環境。\n\n請重新啟動 LibreOffice 以使用語音辨識功能。", "安裝成功", 1)
            return True
                
        except Exception as e:
            logging.error(f"安裝模組時發生錯誤: {str(e)}")
            show_message_box(ctx, f"安裝模組時發生錯誤: {str(e)}", "安裝錯誤", 3)
            return False
        
    except Exception as e:
        error_msg = f"安裝過程中發生錯誤: {str(e)}"
        logging.error(error_msg)
        show_message_box(ctx, error_msg, "安裝錯誤", 3)
        return False