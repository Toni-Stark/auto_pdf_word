import os
import time
import subprocess
import uiautomation as auto
from PIL import ImageGrab
import easyocr
import pyautogui
import numpy as np
import base64
import hashlib
import io
import json
import requests

# 读取配置文件
def load_config():
    """从配置文件读取配置"""
    config = {}
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, 'config.txt')

    if not os.path.exists(config_path):
        print(f"❌ 未找到配置文件: {config_path}")
        return None

    with open(config_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#'):
                if '=' in line:
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip()

                    # 转换数值类型
                    if key in ['CHECK_INTERVAL', 'CAPTURE_INTERVAL']:
                        config[key] = int(value)
                    else:
                        config[key] = value

    return config

# 加载配置
config = load_config()
if not config:
    print("❌ 加载配置失败")
    exit(1)

PDF_FOLDER = config['PDF_FOLDER']
WORD_FOLDER = config['WORD_FOLDER']
WXWORK_WEBHOOK_URL = config['WXWORK_WEBHOOK_URL']
CHECK_INTERVAL = config['CHECK_INTERVAL']
CAPTURE_INTERVAL = config['CAPTURE_INTERVAL']

# 确保文件夹存在
os.makedirs(PDF_FOLDER, exist_ok=True)
os.makedirs(WORD_FOLDER, exist_ok=True)

# 初始化 OCR
print("初始化 OCR 引擎...")
reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)
print("✅ OCR 引擎初始化完成\n")

def check_pdf_files():
    """检查PDF文件夹中是否有PDF文件"""
    if not os.path.exists(PDF_FOLDER):
        return False
    pdf_files = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith('.pdf')]
    return len(pdf_files) > 0

def open_wps_pdf_converter():
    """启动WPS PDF转换"""
    print("→ 启动 WPS PDF 转换...")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    ps1_path = os.path.join(script_dir, 'pdf2word.ps1')

    if not os.path.exists(ps1_path):
        print(f"❌ 未找到PowerShell脚本: {ps1_path}")
        return False

    try:
        subprocess.Popen(['powershell', '-ExecutionPolicy', 'Bypass', '-File', ps1_path])
        print("✅ 已启动 WPS PDF 转换")
        return True
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        return False

    time.sleep(2)  # 等待1秒
def find_wps_window():
    """查找 WPS PDF转换 窗口"""
    wps_win = auto.WindowControl(searchDepth=1, Name='WPS PDF转换')
    return wps_win if wps_win.Exists(0, 0) else None

def handle_password_dialogs():
    """处理密码弹窗"""
    print("\n→ 检测密码弹窗...")
    time.sleep(3)  # 等待3秒让页面加载

    attempt = 0
    max_attempts = 10  # 最多检测10次

    while attempt < max_attempts:
        attempt += 1
        wps_win = find_wps_window()
        if not wps_win:
            print("   ⚠️ 未找到 WPS PDF转换 窗口")
            break

        # 检测"输入密码"文本
        pwd_text = wps_win.TextControl(searchDepth=5, Name='输入密码')
        if pwd_text.Exists(0, 0):
            print(f"   [{attempt}] 检测到密码弹窗")

            # 查找取消按钮
            cancel_btn = wps_win.ButtonControl(searchDepth=5, Name='取消')
            if cancel_btn.Exists(0, 0):
                cancel_btn.Click(simulateMove=False)
                print("   ✓ 已点击取消按钮")
                time.sleep(2)
            else:
                print("   ⚠️ 未找到取消按钮")
                break
        else:
            print("   ✅ 未检测到密码弹窗")
            break

        time.sleep(1)

def find_login_window():
    """查找登录窗口"""
    for window in auto.GetRootControl().GetChildren():
        try:
            window_name = window.Name or ""
            if "登录" in window_name or "WPS账号" in window_name or "扫码" in window_name:
                return window
        except Exception:
            continue
    return None

def capture_and_send_window(window):
    """截取窗口并发送到企业微信"""
    try:
        rect = window.BoundingRectangle
        x1, y1, x2, y2 = rect.left, rect.top, rect.right, rect.bottom
        img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
    except Exception as e:
        print(f"   截图失败: {e}")
        return False

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    raw = buf.getvalue()
    b64 = base64.b64encode(raw).decode("utf-8")
    md5 = hashlib.md5(raw).hexdigest()

    headers = {"Content-Type": "application/json"}
    text_payload = {
        "msgtype": "text",
        "text": {
            "content": (
                f"WPS 需要登录\n"
                f"请扫描下方二维码登录\n"
                f"时间: {time.strftime('%Y-%m-%d %H:%M:%S')}"
            )
        }
    }
    image_payload = {
        "msgtype": "image",
        "image": {"base64": b64, "md5": md5}
    }

    for payload in (text_payload, image_payload):
        try:
            resp = requests.post(
                WXWORK_WEBHOOK_URL,
                data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                headers=headers,
                timeout=15,
            )
            result = resp.json()
            if result.get("errcode") != 0:
                print(f"   企业微信返回错误: {result}")
                return False
        except Exception as e:
            print(f"   发送失败: {e}")
            return False

    return True

def check_and_handle_login():
    """检查并处理登录"""
    print("\n→ 检查登录状态...")
    wps_win = find_wps_window()
    if not wps_win:
        print("   ⚠️ 未找到 WPS PDF转换 窗口")
        return False

    # 查找"请登录"文本或按钮
    login_ctrl = wps_win.TextControl(searchDepth=10, Name='请登录')
    if not login_ctrl.Exists(0, 0):
        login_ctrl = wps_win.ButtonControl(searchDepth=10, Name='请登录')

    if login_ctrl.Exists(0, 0):
        print("   ⚠️ 未登录状态，点击'请登录'...")
        login_ctrl.Click(simulateMove=False)
        print("   ✓ 已点击'请登录'")

        # 等待2秒
        time.sleep(2)

        # 监控登录窗口
        print("   → 开始监控登录窗口...")
        last_capture_time = time.time() - (CAPTURE_INTERVAL - 5)  # 第一次截屏等待5秒

        while True:
            print(f"   [DEBUG] 循环开始 - {time.strftime('%H:%M:%S')}")
            login_window = find_login_window()
            print(f"   [DEBUG] find_login_window() 返回: {login_window is not None}")

            if login_window:
                current_time = time.time()
                # 每60秒截屏一次
                if current_time - last_capture_time >= CAPTURE_INTERVAL:
                    # 点击屏幕中央
                    screen_width, screen_height = pyautogui.size()
                    center_x = screen_width / 2
                    center_y = screen_height / 2
                    print(f"   → 点击屏幕中央 ({center_x:.0f}, {center_y:.0f})...")
                    pyautogui.click(center_x, center_y)
                    time.sleep(0.5)

                    print(f"   [{time.strftime('%H:%M:%S')}] 截屏并发送到企业微信...")
                    if capture_and_send_window(login_window):
                        print("   ✅ 发送成功")
                    else:
                        print("   ⚠️ 发送失败")
                    last_capture_time = current_time
                else:
                    remaining = int(CAPTURE_INTERVAL - (current_time - last_capture_time))
                    print(f"   [{time.strftime('%H:%M:%S')}] 登录窗口存在，{remaining}秒后下次截屏")
            else:
                print(f"   [{time.strftime('%H:%M:%S')}] ✅ 登录窗口已消失")
                break

            print(f"   [DEBUG] 准备 sleep(5)")
            time.sleep(5)
    else:
        print("   ✅ 已登录状态")

    return True

def click_output_range_dropdown():
    """点击输出范围下方的 GroupControl（转换模式选择框）"""
    print("\n→ 点击输出范围下方的选择框...")
    wps_win = find_wps_window()
    if not wps_win:
        print("   ⚠️ 未找到 WPS PDF转换 窗口")
        return False

    # 查找"输出范围"文本控件
    output_range = wps_win.TextControl(searchDepth=15, Name="输出范围")
    if not output_range.Exists(0, 0):
        print("   ⚠️ 未找到'输出范围'文本")
        return False

    print("   ✅ 找到'输出范围'文本")

    # 获取父控件
    parent = output_range.GetParentControl()
    if not parent:
        print("   ⚠️ 未找到父控件")
        return False

    # 获取父控件的所有子控件
    children = parent.GetChildren()

    # 找到"输出范围"在子控件列表中的位置
    found_index = -1
    for i, child in enumerate(children):
        if child.Name == "输出范围":
            found_index = i
            break

    if found_index == -1:
        print("   ⚠️ 在父控件的子控件列表中未找到'输出范围'")
        return False

    # 查找下一个兄弟元素
    if found_index + 1 < len(children):
        next_sibling = children[found_index + 1]

        if next_sibling.ControlTypeName == "GroupControl":
            print("   ✅ 找到输出范围下方的 GroupControl")

            # 获取 GroupControl 的位置
            rect = next_sibling.BoundingRectangle
            center_x = (rect.left + rect.right) / 2
            center_y = (rect.top + rect.bottom) / 2

            # 第一次点击
            print("   → 点击 GroupControl...")
            pyautogui.click(center_x, center_y)
            print("   ✓ 已点击")

            # 等待2秒
            print("   → 等待2秒...")
            time.sleep(2)

            # 点击下方40像素的位置
            click_x = center_x
            click_y = center_y + 40
            print(f"   → 点击下方40像素位置...")
            pyautogui.click(click_x, click_y)
            print("   ✓ 已点击下方位置")

            # 等待2秒
            print("   → 等待2秒...")
            time.sleep(2)

            return True
        else:
            print(f"   ⚠️ 下一个兄弟元素不是 GroupControl")
            return False
    else:
        print("   ⚠️ '输出范围'是最后一个子控件")
        return False

def set_conversion_engine():
    """设置转换引擎为基础版"""
    print("\n→ 设置转换引擎...")
    wps_win = find_wps_window()
    if not wps_win:
        print("   ⚠️ 未找到 WPS PDF转换 窗口")
        return False

    # 查找"转换引擎"文本
    engine_text = wps_win.TextControl(searchDepth=15, Name="转换引擎")
    if not engine_text.Exists(0, 0):
        print("   ⚠️ 未找到'转换引擎'文本")
        return False

    print("   ✅ 找到'转换引擎'")

    # 获取父控件
    parent = engine_text.GetParentControl()
    if not parent:
        print("   ⚠️ 未找到父控件")
        return False

    # 找到"转换引擎"后面的 GroupControl（选择框）
    children = parent.GetChildren()
    found_engine = False
    for child in children:
        if found_engine and child.ControlTypeName == "GroupControl":
            print("   ✅ 找到转换引擎选择框")

            # 获取选择框位置
            rect = child.BoundingRectangle
            x1, y1, x2, y2 = rect.left, rect.top, rect.right, rect.bottom

            # 点击选择框
            print("   → 点击选择框...")
            child.Click(simulateMove=False)
            print("   ✓ 已点击")

            # 等待2秒让下拉菜单展开
            print("   → 等待2秒...")
            time.sleep(2)

            # 截取下拉菜单区域并OCR识别
            dropdown_x1 = x1
            dropdown_y1 = y2
            dropdown_x2 = x2
            dropdown_y2 = y2 + 200

            print(f"   → 截取区域: ({dropdown_x1}, {dropdown_y1}) -> ({dropdown_x2}, {dropdown_y2})")
            img = ImageGrab.grab(bbox=(dropdown_x1, dropdown_y1, dropdown_x2, dropdown_y2))

            # 保存截图用于调试
            debug_img_path = os.path.join(os.path.dirname(__file__), "debug_dropdown.png")
            img.save(debug_img_path)
            print(f"   → 截图已保存: {debug_img_path}")

            print("   → OCR 识别中...")
            results = reader.readtext(np.array(img))

            print(f"   → OCR 识别到 {len(results)} 个文本:")
            for (bbox, text, prob) in results:
                print(f"      - {text} (置信度: {prob:.2f})")

            # 查找"基础版"
            for (bbox, text, prob) in results:
                if "基础版" in text:
                    print(f"   ✅ 找到'基础版' (置信度: {prob:.2f})")

                    # 计算全局坐标
                    center_x = (bbox[0][0] + bbox[2][0]) / 2
                    center_y = (bbox[0][1] + bbox[2][1]) / 2
                    global_x = dropdown_x1 + center_x
                    global_y = dropdown_y1 + center_y

                    # 点击"基础版"
                    print(f"   → 点击坐标: ({global_x:.0f}, {global_y:.0f})")
                    pyautogui.click(global_x, global_y)
                    print("   ✓ 已点击")
                    time.sleep(1)
                    return True

            print("   ⚠️ OCR 未识别到'基础版'")
            return False

        if child.Name == "转换引擎":
            found_engine = True

    print("   ⚠️ 未找到选择框")
    return False

def start_conversion_and_monitor():
    """开始转换并监控状态"""
    print("\n→ 开始转换...")
    wps_win = find_wps_window()
    if not wps_win:
        print("   ⚠️ 未找到 WPS PDF转换 窗口")
        return False

    # 查找"开始转换"文本
    start_text = wps_win.TextControl(searchDepth=15, Name="开始转换")
    if not start_text.Exists(0, 0):
        print("   ⚠️ 未找到'开始转换'按钮")
        return False

    # 获取父控件（ButtonControl）
    start_btn = start_text.GetParentControl()
    print("   ✅ 找到'开始转换'按钮")

    # 点击开始转换
    print("   → 点击'开始转换'...")
    start_btn.Click(simulateMove=False)
    print("   ✓ 已点击")

    # 轮询检测按钮状态
    print("\n→ 监控转换状态...")
    while True:
        time.sleep(5)

        wps_win = find_wps_window()
        if not wps_win:
            print("   ⚠️ WPS 窗口已关闭")
            break

        # 查找"转换中..."文本
        converting_text = wps_win.TextControl(searchDepth=15, Name="转换中...")
        if not converting_text.Exists(0, 0):
            converting_text = wps_win.TextControl(searchDepth=15, Name="转换中")

        if converting_text.Exists(0, 0):
            print(f"   [{time.strftime('%H:%M:%S')}] 转换进行中...")
            continue

        # 查找"开始转换"文本
        start_text = wps_win.TextControl(searchDepth=15, Name="开始转换")
        if start_text.Exists(0, 0):
            print(f"   [{time.strftime('%H:%M:%S')}] ✅ 转换完成!")
            break

    return True

def close_wps():
    """关闭 WPS 进程"""
    print("\n→ 关闭 WPS 进程...")
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'wps.exe'],
                      capture_output=True, text=True)
        print("✅ 已关闭 WPS 进程")
        return True
    except Exception as e:
        print(f"⚠️ 关闭进程失败: {e}")
        return False

def clear_pdf_folder():
    """清空PDF文件夹"""
    print("\n→ 清空PDF文件夹...")
    try:
        pdf_files = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith('.pdf')]
        deleted_count = 0
        for pdf_file in pdf_files:
            pdf_path = os.path.join(PDF_FOLDER, pdf_file)
            try:
                os.remove(pdf_path)
                deleted_count += 1
                print(f"   ✓ 已删除: {pdf_file}")
            except Exception as e:
                print(f"   ⚠️ 删除失败: {pdf_file} - {e}")

        if deleted_count > 0:
            print(f"✅ 成功删除 {deleted_count} 个PDF文件")
        else:
            print("   没有PDF文件需要删除")
        return True
    except Exception as e:
        print(f"⚠️ 清空文件夹失败: {e}")
        return False

def main():
    """主程序"""
    print("=" * 60)
    print("  PDF 自动转换完整流程")
    print("=" * 60)
    print(f"📁 监控文件夹: {PDF_FOLDER}")
    print(f"⏱️  检测间隔: {CHECK_INTERVAL}秒")
    print(f"📌 按 Ctrl+C 停止\n")

    try:
        while True:
            # 步骤1: 监控PDF文件夹
            if check_pdf_files():
                print(f"[{time.strftime('%H:%M:%S')}] 🔍 发现PDF文件\n")

                # 步骤2: 启动WPS PDF转换
                if not open_wps_pdf_converter():
                    print("❌ 启动失败，继续监控...\n")
                    time.sleep(CHECK_INTERVAL)
                    continue

                # 步骤3: 处理密码弹窗
                handle_password_dialogs()

                # 步骤4: 检查并处理登录
                if not check_and_handle_login():
                    print("❌ 登录处理失败")
                    close_wps()
                    time.sleep(CHECK_INTERVAL)
                    continue

                # 步骤4.5: 点击输出范围下方的选择框（转换模式）
                if not click_output_range_dropdown():
                    print("❌ 点击输出范围下方选择框失败")
                    close_wps()
                    time.sleep(CHECK_INTERVAL)
                    continue

                # 步骤5: 设置转换引擎
                if not set_conversion_engine():
                    print("❌ 设置转换引擎失败")
                    close_wps()
                    time.sleep(CHECK_INTERVAL)
                    continue

                # 步骤6: 开始转换并监控
                if not start_conversion_and_monitor():
                    print("❌ 转换失败")
                    close_wps()
                    time.sleep(CHECK_INTERVAL)
                    continue

                # 步骤7: 清理工作
                close_wps()
                clear_pdf_folder()

                print("\n" + "=" * 60)
                print("  ✅ 本次转换流程完成!")
                print("=" * 60)
                print(f"\n[{time.strftime('%H:%M:%S')}] 继续监控...\n")

            time.sleep(CHECK_INTERVAL)

    except KeyboardInterrupt:
        print("\n\n⏹️  程序已停止")
    except Exception as e:
        print(f"\n❌ 程序错误: {e}")

if __name__ == "__main__":
    main()
