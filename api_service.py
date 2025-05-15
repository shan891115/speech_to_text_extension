#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import os
import subprocess
import logging
from pathlib import Path
import platform

def create_api_script(flask_dir, venv_python=None):
    """創建Flask API相關檔案"""
    # 檢測作業系統
    system = platform.system()
    is_windows = system == "Windows"
    is_macos = system == "Darwin"
    is_linux = system == "Linux"
    
    # 創建Flask API Python檔案
    api_file = flask_dir / 'speech_api.py'
    
    # 獲取虛擬環境 Python 路徑（如果提供）
    venv_python_path = f"# 使用虛擬環境 Python: {venv_python}" if venv_python else ""
    
    # 根據不同作業系統創建啟動腳本
    if is_windows:
        # Windows 啟動腳本 (.bat)
        start_script = flask_dir / 'start_api.bat'
        with open(start_script, 'w', encoding='utf-8') as f:
            if venv_python:
                f.write(f'@echo off\necho 啟動語音辨識 API 服務...\n"{venv_python}" "{flask_dir / "speech_api.py"}"\npause')
            else:
                f.write(f'@echo off\necho 啟動語音辨識 API 服務...\npython "{flask_dir / "speech_api.py"}"\npause')
    else:
        # Linux/Mac 啟動腳本 (.sh)
        start_script = flask_dir / 'start_api.sh'
        with open(start_script, 'w', encoding='utf-8') as f:
            if venv_python:
                f.write(f'#!/bin/bash\necho "啟動語音辨識 API 服務..."\n"{venv_python}" "{flask_dir / "speech_api.py"}"\nread -p "按任意鍵繼續..."')
            else:
                f.write(f'#!/bin/bash\necho "啟動語音辨識 API 服務..."\npython3 "{flask_dir / "speech_api.py"}"\nread -p "按任意鍵繼續..."')
        
        # 設置執行權限
        try:
            os.chmod(start_script, 0o755)
        except:
            logging.warning(f"無法設置 {start_script} 的執行權限")
    
    # 創建 API 腳本文件
    with open(api_file, 'w', encoding='utf-8') as f:
        f.write(rf'''#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import platform
import logging
from datetime import datetime
from pathlib import Path

{venv_python_path}

# 設定日誌
log_dir = Path.home() / '.libreoffice' / 'speech_to_text_logs'
log_dir.mkdir(parents=True, exist_ok=True)
log_file = log_dir / f'speech_api_{{datetime.now().strftime("%Y%m%d")}}.log'
logging.basicConfig(
    filename=str(log_file), 
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# 記錄當前 Python 路徑與版本
logging.debug(f"Python 執行檔: {{sys.executable}}")
logging.debug(f"Python 版本: {{sys.version}}")
logging.debug(f"作業系統: {{platform.system()}} {{platform.release()}}")

# 動態添加家目錄下的 Python 模組路徑
home_dir = Path.home()
python_lib_dirs = []

# 識別當前 Python 主要版本
python_major = sys.version_info.major
python_minor = sys.version_info.minor

# 優先添加用戶虛擬環境路徑
system = platform.system()
if system == 'Windows':
    venv_site_packages = home_dir / '.libreoffice' / 'python_env' / 'venv' / 'Lib' / 'site-packages'
else:
    # Linux/Mac 虛擬環境結構
    venv_site_packages = home_dir / '.libreoffice' / 'python_env' / 'venv' / 'lib' / f'python{{python_major}}.{{python_minor}}' / 'site-packages'

if venv_site_packages.exists():
    python_lib_dirs.append(venv_site_packages)
    logging.debug(f"添加虛擬環境路徑: {{venv_site_packages}}")

# 根據作業系統添加適當的路徑
if system == 'Windows':
    # Windows 路徑模式
    for minor_ver in range(python_minor, python_minor + 3):  # 當前版本及未來兩個小版本
        python_lib_dirs.append(
            home_dir / 'AppData' / 'Local' / 'Programs' / 'Python' / 
            f'Python{{python_major}}{{minor_ver}}' / 'Lib' / 'site-packages'
        )
    # 添加 pip 安裝的用戶模組路徑
    python_lib_dirs.append(
        home_dir / 'AppData' / 'Roaming' / 'Python' / 
        f'Python{{python_major}}{{python_minor}}' / 'site-packages'
    )
elif system == 'Darwin':
    # macOS 路徑模式
    for minor_ver in range(python_minor - 1, python_minor + 2):  # 前一個版本、當前版本和下一個版本
        if minor_ver > 0:  # 確保次要版本號有效
            python_lib_dirs.append(
                home_dir / 'Library' / 'Python' / f'{{python_major}}.{{minor_ver}}' / 'lib' / 'python' / 'site-packages'
            )
    # 常見的 homebrew Python 路徑
    python_lib_dirs.append(Path(f'/usr/local/lib/python{{python_major}}.{{python_minor}}/site-packages'))
    python_lib_dirs.append(Path(f'/opt/homebrew/lib/python{{python_major}}.{{python_minor}}/site-packages'))
else:
    # Linux 路徑模式
    for minor_ver in range(python_minor - 1, python_minor + 2):  # 前一個版本、當前版本和下一個版本
        if minor_ver > 0:  # 確保次要版本號有效
            python_lib_dirs.append(
                home_dir / '.local' / 'lib' / f'python{{python_major}}.{{minor_ver}}' / 'site-packages'
            )
    # 常見的系統路徑
    python_lib_dirs.append(Path(f'/usr/lib/python{{python_major}}.{{python_minor}}/site-packages'))
    python_lib_dirs.append(Path(f'/usr/lib/python{{python_major}}/dist-packages'))
    python_lib_dirs.append(Path(f'/usr/local/lib/python{{python_major}}.{{python_minor}}/site-packages'))

# 添加自定義模組目錄
python_lib_dirs.append(home_dir / '.libreoffice' / 'python_modules')

# 添加所有可能的路徑
for path in python_lib_dirs:
    if path.exists() and str(path) not in sys.path:
        sys.path.append(str(path))
        logging.debug(f"添加路徑: {{path}}")

# 嘗試導入必要模組
required_modules = {{
    'flask': 'Flask Web 框架',
    'speech_recognition': 'SpeechRecognition 語音辨識',
    'pyaudio': 'PyAudio 音訊處理'
}}

missing_modules = []

for module_name, description in required_modules.items():
    try:
        module = __import__(module_name)
        if hasattr(module, '__version__'):
            logging.debug(f"{{description}} 版本: {{module.__version__}}")
        else:
            logging.debug(f"{{description}} 已導入")
    except ImportError as e:
        logging.error(f"無法導入 {{description}}: {{e}}")
        missing_modules.append(module_name)

if missing_modules:
    error_msg = f"缺少必要模組: {{', '.join(missing_modules)}}，請安裝所需模組。"
    logging.error(error_msg)
    print(error_msg)
    sys.exit(1)

# Flask 應用程式
try:
    from flask import Flask, request, jsonify
    import speech_recognition as sr
    
    app = Flask(__name__)

    def check_microphone():
        """檢查麥克風是否可用"""
        try:
            mic_list = sr.Microphone.list_microphone_names()
            if not mic_list:
                logging.warning("未檢測到麥克風裝置")
                return False
            logging.debug(f"檢測到 {{len(mic_list)}} 個麥克風裝置")
            return True
        except Exception as e:
            logging.error(f"檢查麥克風時發生錯誤: {{e}}")
            return False

    @app.route('/', methods=['GET'])
    def index():
        """API 根路徑，返回服務狀態"""
        mic_available = check_microphone()
        return jsonify({{
            "status": "running",
            "microphone_available": mic_available,
            "python_version": sys.version,
            "timestamp": datetime.now().isoformat()
        }})

    @app.route('/mic_check', methods=['GET'])
    def mic_check():
        """檢查麥克風可用性"""
        mic_available = check_microphone()
        if mic_available:
            return jsonify({{"success": True, "message": "麥克風可用"}})
        else:
            return jsonify({{"success": False, "error": "未檢測到可用麥克風"}})

    @app.route('/recognize', methods=['POST'])
    def recognize_speech():
        """語音辨識端點"""
        try:
            # 先檢查麥克風
            if not check_microphone():
                return jsonify({{"success": False, "error": "未檢測到可用麥克風"}})
                
            # 獲取設置參數
            language = request.json.get('language', 'zh-TW')
            timeout = request.json.get('timeout', 10)  # 增加默認等待時間到10秒
            phrase_time_limit = request.json.get('phrase_time_limit', None)
            # 新增靜音等待參數 - 在檢測到語音停止後，再等待這麼久才結束識別
            pause_threshold = request.json.get('pause_threshold', 5.0)  # 默認等待5秒無聲才結束
            non_speaking_duration = request.json.get('non_speaking_duration', 1.0)  # 設定檢測無聲時間閾值
            
            logging.debug(f"開始辨識，語言: {{language}}, 超時: {{timeout}}秒, 靜音等待: {{pause_threshold}}秒")
            
            recognizer = sr.Recognizer()
            # 設定靜音等待時間 - 檢測到停止說話後再等多久才算結束
            recognizer.pause_threshold = float(pause_threshold)
            # 設定檢測靜音時間閾值
            recognizer.non_speaking_duration = float(non_speaking_duration)
            
            try:
                with sr.Microphone() as source:
                    logging.debug("麥克風開啟")
                    # 調整環境噪音
                    recognizer.adjust_for_ambient_noise(source)
                    logging.debug("已調整環境噪音")
                
                    # 提示用戶開始說話
                    print("請開始說話...")
                    
                    # 取得語音輸入 - 注意使用更長的timeout和phrase_time_limit
                    audio = recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
                    logging.debug("已擷取音訊")
            except sr.WaitTimeoutError:
                return jsonify({{"success": False, "error": "聆聽超時，未檢測到語音"}})
            except Exception as e:
                logging.error(f"麥克風使用錯誤: {{str(e)}}")
                return jsonify({{"success": False, "error": f"麥克風使用錯誤: {{str(e)}}"}})
            
            try:
                # 使用 Google API 辨識語音
                recognized_text = recognizer.recognize_google(audio, language=language)
                logging.debug(f"辨識結果: {{recognized_text}}")
                
                return jsonify({{
                    "success": True, 
                    "text": recognized_text,
                    "language": language
                }})
            except sr.UnknownValueError:
                logging.warning("無法辨識語音內容")
                return jsonify({{"success": False, "error": "無法辨識語音內容"}})
            except sr.RequestError as e:
                logging.error(f"Google API 請求錯誤: {{str(e)}}")
                return jsonify({{"success": False, "error": f"語音辨識服務錯誤: {{str(e)}}"}})
            
        except Exception as e:
            logging.error(f"處理請求時發生錯誤: {{str(e)}}")
            return jsonify({{"success": False, "error": f"發生錯誤: {{str(e)}}"}})

    @app.errorhandler(404)
    def not_found(error):
        """處理 404 錯誤"""
        return jsonify({{"success": False, "error": "端點不存在"}}), 404

    @app.errorhandler(500)
    def server_error(error):
        """處理 500 錯誤"""
        logging.error(f"伺服器錯誤: {{str(error)}}")
        return jsonify({{"success": False, "error": "伺服器內部錯誤"}}), 500
                
except Exception as e:
    logging.error(f"初始化 Flask 或 SpeechRecognition 發生錯誤: {{e}}")
    print(f"初始化錯誤: {{e}}")
    sys.exit(1)                

if __name__ == '__main__':
    # 檢查端口是否已被使用
    import socket
    port = 5000
    retry = 0
    
    while retry < 3:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        result = sock.connect_ex(('127.0.0.1', port))
        sock.close()
        
        if result == 0:  # 端口已被使用
            logging.warning(f"端口 {{port}} 已被占用，嘗試使用 {{port+1}}")
            port += 1
            retry += 1
        else:
            break
    
    if retry >= 3:
        logging.error(f"無法找到可用端口，請確認 5000-5002 端口是否被占用")
        print("錯誤: 無法找到可用端口，請確認 5000-5002 端口是否被占用")
        sys.exit(1)
        
    print(f"啟動語音辨識 API 服務在 http://127.0.0.1:{{port}}")
    logging.info(f"啟動語音辨識 API 服務在 http://127.0.0.1:{{port}}")
    app.run(host='127.0.0.1', port=port, debug=False)
''')

    # 設置API腳本的執行權限 (Linux/Mac)
    if not is_windows:
        try:
            os.chmod(api_file, 0o755)
        except:
            logging.warning(f"無法設置 {api_file} 的執行權限")

def start_api_server():
    """直接在Python中啟動語音辨識API服務"""
    try:
        from pathlib import Path
        import subprocess
        import os
        import platform
        
        # 檢測作業系統
        system = platform.system()
        is_windows = system == "Windows"
        is_macos = system == "Darwin"
        is_linux = system == "Linux"
        
        home_path = Path.home()
        flask_dir = home_path / '.libreoffice' / 'speech_api'
        api_script = flask_dir / 'speech_api.py'
        
        # 根據不同作業系統尋找虛擬環境中的 Python 路徑
        if is_windows:
            venv_python = home_path / '.libreoffice' / 'python_env' / 'venv' / 'Scripts' / 'python.exe'
        else:  # Linux/Mac
            venv_python = home_path / '.libreoffice' / 'python_env' / 'venv' / 'bin' / 'python'
        
        # 決定用哪個 Python 啟動 API 服務
        python_executable = str(venv_python) if venv_python.exists() else sys.executable
        logging.debug(f"使用 Python 執行檔: {python_executable}")
        
        # 首先檢查必要目錄和文件是否存在
        if not flask_dir.exists():
            logging.debug(f"創建 API 目錄 {flask_dir}")
            flask_dir.mkdir(parents=True, exist_ok=True)
        
        if not api_script.exists():
            logging.debug("API 腳本不存在，創建新的腳本")
            # 如果虛擬環境存在，傳遞給 create_api_script
            if venv_python.exists():
                create_api_script(flask_dir, venv_python)
            else:
                create_api_script(flask_dir)
        
        logging.debug(f"啟動 API 腳本：{api_script}")
        
        # 確認腳本存在後再啟動
        if not api_script.exists():
            logging.error(f"API 腳本不存在：{api_script}")
            return False
        
        # 確保權限正確 (僅 Linux/Mac)
        if not is_windows:
            try:
                os.chmod(api_script, 0o755)
            except:
                logging.warning(f"無法設置 {api_script} 的執行權限，但仍會嘗試執行")
        
        # 使用選定的 Python 啟動 API 服務
        process = subprocess.Popen(
            [python_executable, str(api_script)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False,  # 避免shell=True的路徑問題
            text=True     # 讓stdout/stderr為文本模式
        )
        
        # 等待一小段時間確認進程啟動正常
        import time
        time.sleep(1)
        
        # 檢查進程是否立即退出
        if process.poll() is not None:
            # 進程已退出，可能有錯誤
            stdout, stderr = process.communicate()
            logging.error(f"API進程立即退出，返回碼：{process.returncode}")
            logging.error(f"標準輸出：{stdout}")
            logging.error(f"錯誤輸出：{stderr}")
            
            # 如果使用虛擬環境失敗，嘗試使用系統 Python
            if python_executable == str(venv_python):
                logging.debug("虛擬環境 Python 啟動失敗，嘗試使用系統 Python")
                return start_api_server_with_system_python(api_script)
            return False
        
        # 給API一些啟動時間
        time.sleep(3)
        
        # 嘗試連接API確認是否啟動
        import socket
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        result = sock.connect_ex(('127.0.0.1', 5000))
        sock.close()
        
        if result == 0:  # 端口可連接
            logging.info("API 服務啟動成功")
            return True
        else:
            # 嘗試備用端口
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            result = sock.connect_ex(('127.0.0.1', 5001))
            sock.close()
            
            if result == 0:  # 備用端口可連接
                logging.info("API 服務在備用端口 5001 啟動成功")
                return True
            else:
                logging.error("無法連接到 API 服務")
                # 如果使用虛擬環境失敗，嘗試使用系統 Python
                if python_executable == str(venv_python):
                    logging.debug("虛擬環境 Python 啟動失敗，嘗試使用系統 Python")
                    return start_api_server_with_system_python(api_script)
                return False
                
    except Exception as e:
        logging.error(f"啟動API服務失敗: {e}")
        return False

def start_api_server_with_system_python(api_script):
    """使用系統 Python 啟動 API 服務 (備用方案)"""
    try:
        import subprocess
        import time
        import platform
        
        system = platform.system()
        is_windows = system == "Windows"
        
        logging.debug(f"使用系統 Python 啟動 API: {sys.executable}")
        
        # 使用系統 Python 啟動 API
        process = subprocess.Popen(
            [sys.executable, str(api_script)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False,
            text=True
        )
        
        # 等待確認進程啟動
        time.sleep(1)
        if process.poll() is not None:
            stdout, stderr = process.communicate()
            logging.error(f"系統 Python API進程立即退出，返回碼：{process.returncode}")
            logging.error(f"標準輸出：{stdout}")
            logging.error(f"錯誤輸出：{stderr}")
            
            # 在 Linux/Mac 上，嘗試使用 python3 命令
            if not is_windows:
                try:
                    logging.debug("嘗試使用 python3 命令啟動 API")
                    alt_process = subprocess.Popen(
                        ["python3", str(api_script)],
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        shell=False,
                        text=True
                    )
                    time.sleep(1)
                    if alt_process.poll() is None:
                        # 給API一些啟動時間
                        time.sleep(3)
                        
                        # 嘗試連接API確認是否啟動
                        import socket
                        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                        result = sock.connect_ex(('127.0.0.1', 5000))
                        sock.close()
                        
                        if result == 0 or sock.connect_ex(('127.0.0.1', 5001)) == 0:
                            logging.info("使用 python3 命令啟動 API 服務成功")
                            return True
                except:
                    logging.error("使用 python3 命令啟動 API 失敗")
            
            return False
        
        # 給API一些啟動時間
        time.sleep(3)
        
        # 嘗試連接API確認是否啟動
        import socket
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        result = sock.connect_ex(('127.0.0.1', 5000))
        sock.close()
        
        if result == 0 or sock.connect_ex(('127.0.0.1', 5001)) == 0:
            logging.info("使用系統 Python 啟動 API 服務成功")
            return True
        else:
            logging.error("使用系統 Python 啟動 API 服務失敗")
            return False
            
    except Exception as e:
        logging.error(f"使用系統 Python 啟動 API 服務失敗: {e}")
        return False