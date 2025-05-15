import sys
import unohelper
import officehelper
import logging
import os
import tempfile
import importlib.util
import platform
from pathlib import Path

from com.sun.star.task import XJobExecutor
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL

sys.path.append(os.path.dirname(__file__))

# 導入自定義模組
from utils import setup_logging, check_module_installed, show_message_box
from module_installer import install_modules_directly, fix_venv_permissions
from api_service import create_api_script, start_api_server, start_api_server_with_system_python

class SpeechToTextJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx
        self.log_dir = setup_logging()
        logging.debug("SpeechToTextJob initialized")
        
        try:
            self.sm = ctx.getServiceManager()
            try:
                self.desktop = XSCRIPTCONTEXT.getDesktop()
            except NameError:
                self.desktop = self.sm.createInstanceWithContext(
                    "com.sun.star.frame.Desktop", self.ctx
                )
            logging.debug("Desktop service initialized")
            
            # 檢查首次安裝
            self.check_first_install()
            
        except Exception as e:
            logging.error(f"Initialization error: {e}")

    def check_first_install(self):
        try:
            from pathlib import Path
            home_path = Path.home()
            config_marker = home_path / '.libreoffice' / 'speech_to_text_installed'
            
            # 创建一个简单的HTTP客户端，避免导入requests模块
            import sys
            
            # 如果requests已经在sys.modules中，先保存一下它的状态以便后续恢复
            has_real_requests = 'requests' in sys.modules
            old_requests = sys.modules.get('requests', None)
            
            # 创建一个简单HTTP客户端类，替代requests模块
            import http.client
            import json
            import urllib.parse
            
            class SimpleHTTPClient:
                @staticmethod
                def get(url):
                    parsed = urllib.parse.urlparse(url)
                    conn = http.client.HTTPConnection(parsed.netloc)
                    conn.request("GET", parsed.path or '/')
                    resp = conn.getresponse()
                    body = resp.read()
                    conn.close()
                    return SimpleResponse(resp.status, resp.reason, body)
            
                @staticmethod
                def post(url, json_data=None):
                    parsed = urllib.parse.urlparse(url)
                    headers = {"Content-Type": "application/json"}
                    
                    body = None
                    if json_data is not None:
                        try:
                            body = json.dumps(json_data).encode('utf-8')
                        except Exception as e:
                            logging.error(f"JSON 序列化錯誤: {e}")
                            raise
                
                    conn = http.client.HTTPConnection(parsed.netloc)
                    conn.request("POST", parsed.path or '/', body=body, headers=headers)
                    resp = conn.getresponse()
                    body = resp.read()
                    conn.close()
                    return SimpleResponse(resp.status, resp.reason, body)
        
            class SimpleResponse:
                def __init__(self, status_code, reason, content):
                    self.status_code = status_code
                    self.reason = reason
                    self.content = content
            
                def json(self):
                    if isinstance(self.content, bytes):
                        return json.loads(self.content.decode('utf-8'))
                    else:
                        return json.loads(self.content)
            
                def __str__(self):
                    return f"Response [status_code={self.status_code}]"
            
            # 创建一个模拟的spec对象
            class MockSpec:
                def __init__(self):
                    self.name = "requests"
                    self.loader = None
                    self.origin = None
                    self.submodule_search_locations = None
                    
            # 创建虚拟requests模块
            mock_spec = MockSpec()
            virtual_requests = type('requests', (), {
                'get': SimpleHTTPClient.get,
                'post': SimpleHTTPClient.post,
                'exceptions': type('exceptions', (), {
                    'ConnectionError': Exception,
                    'RequestException': Exception
                }),
                '__spec__': mock_spec  # 添加__spec__属性避免报错
            })
            
            # 暂时替换requests模块
            sys.modules['requests'] = virtual_requests
            
            # 现在可以安全地检查模块安装情况
            # 检查必要模块 (不论是否首次安装都检查)
            sr_installed = check_module_installed('speech_recognition')
            pyaudio_installed = check_module_installed('pyaudio')
            flask_installed = check_module_installed('flask')
            requests_installed = check_module_installed('requests')
              # 恢复原始的requests模块（如果存在）
            if has_real_requests and old_requests is not None:
                sys.modules['requests'] = old_requests
            elif not has_real_requests:
                # 如果之前没有真实的requests模块，则删除我们添加的虚拟模块
                if 'requests' in sys.modules:
                    del sys.modules['requests']
            
            # 在 Linux 上，修復可能的權限問題
            if platform.system() == "Linux":
                logging.debug("檢測到 Linux 系統，嘗試修復虛擬環境權限")
                fix_venv_permissions()
                # 修復權限後重新檢查模組
                sr_installed = check_module_installed('speech_recognition')
                pyaudio_installed = check_module_installed('pyaudio')
                flask_installed = check_module_installed('flask')
                requests_installed = check_module_installed('requests')
            
            # 其余代码保持不变
            missing_modules = []
            if not sr_installed:
                missing_modules.append("speech_recognition")
            if not pyaudio_installed:
                missing_modules.append("pyaudio")
            if not flask_installed:
                missing_modules.append("flask")
            if not requests_installed:
                missing_modules.append("requests")
            
            # 無論是否首次安裝，如有缺少模組都提示安裝
            if missing_modules:
                message = f"語音辨識擴充套件缺少以下必要模組：\n{', '.join(missing_modules)}\n\n要現在安裝嗎？"
                result = show_message_box(self.ctx, message, "安裝必要模組", QUERYBOX, BUTTONS_YES_NO)
                
                if result == YES:
                    # 直接在Python中安裝模組，無需批次檔
                    if install_modules_directly(self.ctx):
                        # 安裝成功後，嘗試立即啟動 API 服務
                        logging.debug("模組安裝成功，嘗試啟動 API 服務")
                        if start_api_server():
                            logging.debug("API 服務已自動啟動")
                        return
                    else:
                        # 安裝失敗
                        show_message_box(self.ctx, "缺少必要模組，語音辨識功能可能無法正常運作。", "警告", WARNINGBOX)
                        return
                else:
                    # 如果用戶選擇不安裝，顯示警告
                    show_message_box(self.ctx, "缺少必要模組，語音辨識功能可能無法正常運作。", "警告", WARNINGBOX)
                    return
        
            # 如果標記檔不存在，則視為首次安裝 (但已確認模組都已安裝)
            if not config_marker.exists():
                logging.debug("First install detected, all modules are installed")
                
                # 創建 Flask API 及必要檔案
                flask_dir = home_path / '.libreoffice' / 'speech_api'
                flask_dir.mkdir(parents=True, exist_ok=True)
                
                # 尋找虛擬環境中的 Python
                venv_python = home_path / '.libreoffice' / 'python_env' / 'venv' / 'Scripts' / 'python.exe'
                
                # 創建API腳本文件，如果存在虛擬環境則使用它
                if venv_python.exists():
                    create_api_script(flask_dir, venv_python)
                else:
                    create_api_script(flask_dir)
                
                # 創建標記檔，表示已安裝完成
                config_marker.parent.mkdir(parents=True, exist_ok=True)
                with open(config_marker, 'w', encoding='utf-8') as f:
                    f.write("Installed")
                
                # 自動啟動 API 服務
                if start_api_server():
                    logging.debug("API 服務已自動啟動")
                    show_message_box(self.ctx, "語音辨識功能已設置完成並啟動服務。您可以開始使用該功能了。", "安裝成功", INFOBOX)
                else:
                    show_message_box(self.ctx, "語音辨識功能已設置完成，但無法自動啟動服務。點擊語音辨識按鈕時將再次嘗試啟動。", "安裝成功", INFOBOX)
            else:
                # 即使不是首次安裝，也嘗試啟動 API 服務 (如果服務尚未運行)
                try:
                    import http.client, urllib.parse
                    conn = http.client.HTTPConnection("127.0.0.1", 5000)
                    conn.request("GET", "/")
                    resp = conn.getresponse()
                    
                    # 如果能夠連接到服務，則不需要啟動
                    if resp.status == 200:
                        logging.debug("API 服務已經運行中")
                        return
                        
                    conn.close()
                except:
                    # 服務未運行，嘗試啟動
                    logging.debug("嘗試自動啟動 API 服務")
                    if start_api_server():
                        logging.debug("API 服務已自動啟動")
                    else:
                        logging.debug("無法自動啟動 API 服務，將等待用戶手動啟动")
                
        except Exception as e:
            logging.error(f"Error in check_first_install: {e}")
            # 不顯示詳細錯誤，只顯示通用消息，避免嚇到用戶
            # show_message_box(self.ctx, f"安裝檢查過程中發生錯誤：{str(e)}", "錯誤", ERRORBOX)

    def trigger(self, args):
        logging.debug(f"Trigger called with args: {args}")
        try:
            # 添加嵌入式HTTP客戶端代碼 - 不需要requests模組
            import sys
            import http.client
            import json
            import urllib.parse
        
            # 定義一個簡單的HTTP客戶端類，替代requests的基本功能
            class SimpleHTTPClient:
                @staticmethod
                def get(url):
                    parsed = urllib.parse.urlparse(url)
                    conn = http.client.HTTPConnection(parsed.netloc)
                    conn.request("GET", parsed.path or '/')
                    resp = conn.getresponse()
                    body = resp.read()
                    conn.close()
                    return SimpleResponse(resp.status, resp.reason, body)
            
                @staticmethod
                def post(url, json_data=None):
                    parsed = urllib.parse.urlparse(url)
                    headers = {"Content-Type": "application/json"}
                    
                    # 修复 JSON 序列化問題
                    body = None
                    if json_data is not None:
                        try:
                            body = json.dumps(json_data).encode('utf-8')
                        except Exception as e:
                            logging.error(f"JSON 序列化錯誤: {e}")
                            raise
                
                    conn = http.client.HTTPConnection(parsed.netloc)
                    conn.request("POST", parsed.path or '/', body=body, headers=headers)
                    resp = conn.getresponse()
                    body = resp.read()
                    conn.close()
                    return SimpleResponse(resp.status, resp.reason, body)
        
            class SimpleResponse:
                def __init__(self, status_code, reason, content):
                    self.status_code = status_code
                    self.reason = reason
                    self.content = content
            
                def json(self):
                    # 確保 content 是字節類型，并正確解碼
                    if isinstance(self.content, bytes):
                        return json.loads(self.content.decode('utf-8'))
                    else:
                        return json.loads(self.content)
            
                def __str__(self):
                    return f"Response [status_code={self.status_code}]"
        
            # 添加到模組名稱空間，這樣其他方法可以使用
            sys.modules['requests'] = type('requests', (), {
                'get': SimpleHTTPClient.get,
                'post': SimpleHTTPClient.post,
                'exceptions': type('exceptions', (), {
                    'ConnectionError': Exception,
                    'RequestException': Exception
                })
            })
        
            # 現在可以從sys.modules獲取"requests"
            import requests
            logging.debug("Successfully created fallback requests module")
        
            # 檢查API服務是否啟動
            self.ensure_api_running()
        
            # 執行語音辨識
            self.start_speech_to_text()
        except Exception as e:
            logging.error(f"Error in trigger: {e}")
            show_message_box(self.ctx, f"執行語音辨識時發生錯誤：{str(e)}", "語音辨識錯誤", ERRORBOX)
    
    def ensure_api_running(self):
        """確保API服務正在運行"""
        import requests
        try:
            # 嘗試連接API
            requests.get("http://127.0.0.1:5000")
            logging.debug("API service is running")
        except Exception:  # 不再區分具體異常類型
            # 檢查是否已安裝為Windows服務
            try:
                # 檢查服務狀態
                service_status = subprocess.call(['sc', 'query', 'SpeechRecognitionAPI'], 
                                     stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                is_service_installed = service_status == 0
            except:
                is_service_installed = False

            if is_service_installed:
                # 嘗試啟動服務
                message = "語音辨識服務未運行。是否啟動Windows服務？"
                result = show_message_box(self.ctx, message, "啟動語音辨識服務", QUERYBOX, BUTTONS_YES_NO)
        
                if result == YES:
                    try:
                        subprocess.call(['net', 'start', 'SpeechRecognitionAPI'])
                        # 給API一些啟動時間
                        import time
                        time.sleep(3)
                        # 再次檢查API是否運行，但不做異常處理
                        requests.get("http://127.0.0.1:5000")
                    except Exception as e:
                        logging.error(f"啟動服務失敗: {e}")
                        # 如果服務啟動失敗，嘗試直接啟動API
                        message = "服務啟動失敗，是否直接啟动API服務？"
                        result = show_message_box(self.ctx, message, "啟動API服務", QUERYBOX, BUTTONS_YES_NO)
                        if result == YES:
                            if start_api_server():
                                # 再次檢查API是否運行
                                try:
                                    requests.get("http://127.0.0.1:5000")
                                    logging.debug("API服務啟動成功")
                                except Exception:
                                    raise Exception("API服務啟動失敗")
                            else:
                                raise Exception("API服務啟動失敗")
                        else:
                            raise Exception("無法使用語音辨識，API服務未運行")
                else:
                    raise Exception("無法使用語音辨識，API服務未運行")
            else:
                # 直接使用Python啟動API
                message = "語音辨識API服務未運行。是否啟动服務？"
                result = show_message_box(self.ctx, message, "啟動語音辨識服務", QUERYBOX, BUTTONS_YES_NO)
        
                if result == YES:
                    # 使用Python直接啟動API，避免使用批次檔
                    if start_api_server():
                        # 給API一些啟動時間
                        import time
                        time.sleep(3)
                    
                        # 再次檢查API是否運行
                        try:
                            requests.get("http://127.0.0.1:5000")
                            logging.debug("API服務啟動成功")
                        except Exception:
                            # 如果還是無法連接，顯示錯誤
                            message = "API服務啟動失敗。請檢查Python環境是否正確設置，以及必要模組是否已安裝。"
                            show_message_box(self.ctx, message, "API服務啟動失敗", ERRORBOX)
                            raise Exception("API服務啟動失敗")
                    else:
                        message = "API服務啟動失敗。請檢查Python環境是否正確設置，以及必要模組是否已安裝。"
                        show_message_box(self.ctx, message, "API服務啟動失敗", ERRORBOX)
                        raise Exception("API服務啟動失敗")
                else:
                    raise Exception("無法使用語音辨識，API服務未運行")

    def start_speech_to_text(self):
        logging.debug("Starting speech to text")
        try:
            # 獲取當前文件
            model = self.desktop.getCurrentComponent()
        
            # 確保是文字文件
            if not hasattr(model, "Text"):
                logging.error("Current component is not a text document")
                show_message_box(self.ctx, "請在文字文件中使用此功能", "語音辨識錯誤", ERRORBOX)
                return
            
            text = model.Text
            
            # 获取当前光标位置
            controller = model.getCurrentController()
            view_cursor = controller.getViewCursor()
            
            # 显示正在聆听消息（创建临时cursor用于此消息）
            temp_cursor = text.createTextCursorByRange(view_cursor.getStart())
            listening_text = "正在聆聽..."
            text.insertString(temp_cursor, listening_text, 0)
        
            try:
                # 呼叫API進行語音辨識 - 使用更长的参数值
                import requests
                response = requests.post(
                    "http://127.0.0.1:5000/recognize",
                    json_data={
                        "language": "zh-TW", 
                        "timeout": 15,         # 增加到15秒
                        "pause_threshold": 8.0, # 增加到8秒无声才结束
                        "non_speaking_duration": 1.5  # 增加检测静音阈值
                    }
                )
                
                # 删除"正在聆听..."消息 - 修复删除操作
                try:
                    # 创建新的光标范围来选择之前插入的文本
                    delete_cursor = text.createTextCursorByRange(view_cursor.getStart())
                    # 移动光标选择"正在聆聽..."
                    delete_cursor.goLeft(len(listening_text), True)
                    # 删除选中的文本
                    delete_cursor.setString("")
                except Exception as delete_error:
                    logging.error(f"删除临时文本错误: {delete_error}")
                    # 如果删除失败，继续处理识别结果
                
                if response.status_code == 200:
                    result = response.json()
                    if result.get("success"):
                        recognized_text = result.get("text", "")
                        # 在当前光标位置插入识别结果
                        text.insertString(view_cursor, recognized_text, 0)
                    else:
                        error_msg = result.get("error", "未知錯誤")
                        show_message_box(self.ctx, f"辨識失敗：{error_msg}", "語音辨識錯誤", WARNINGBOX)
                else:
                    show_message_box(self.ctx, f"API服務錯誤：HTTP狀態碼 {response.status_code}", "語音辨識錯誤", ERRORBOX)
                
            except Exception as e:
                # 删除"正在聆听..."消息 - 在异常处理中也使用相同的修复方法
                try:
                    delete_cursor = text.createTextCursorByRange(view_cursor.getStart())
                    delete_cursor.goLeft(len(listening_text), True)
                    delete_cursor.setString("")
                except:
                    logging.error("无法删除临时提示文本")
                
                show_message_box(self.ctx, f"API請求錯誤：{str(e)}", "語音辨識錯誤", ERRORBOX)
                logging.error(f"API request error: {e}")
        
        except Exception as e:
            logging.error(f"Error in start_speech_to_text: {e}")
            show_message_box(self.ctx, f"語音辨識發生錯誤：{str(e)}", "語音辨識錯誤", ERRORBOX)

# Starting from Python IDE
def main():
    logging.debug("Main function started")
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        logging.debug("Got context from XSCRIPTCONTEXT")
    except NameError:
        logging.debug("Getting context from bootstrap")
        ctx = officehelper.bootstrap()
        if ctx is None:
            logging.error("Could not bootstrap default Office")
            print("ERROR: Could not bootstrap default Office.")
            sys.exit(1)
    
    job = SpeechToTextJob(ctx)
    job.trigger("StartSpeechRecognition")  # 使用與 XML 中相同的命令

# Starting from command line
if __name__ == "__main__":
    main()

# 註冊 UNO 組件
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    SpeechToTextJob,
    "org.extension.speech.to.text.do",  # 確保與 XML 完全匹配
    ("com.sun.star.task.Job",),
)