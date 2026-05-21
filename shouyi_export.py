"""
收益数据导出脚本：从专区收益统计页面下载数据
"""
from playwright.sync_api import sync_playwright
import os
import yaml
import sys
from datetime import datetime
from feishu_notify import send_image as feishu_send_image

sys.stdout.reconfigure(encoding='utf-8')

# 从环境变量读取目标 chat_id（agent 传入）
GROUP_CHAT_ID = os.environ.get("FEISHU_NOTIFY_CHAT_ID", "")


def load_config():
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.yaml'), 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def download_revenue(start_date, end_date, url="https://dui.epoint.com.cn/transferplatform/pages/transferplatform/yfw/strategicmaplist",
                     download_button_selector="#dataexport .mini-button-text", debug_port=9222):
    """
    使用浏览器自动化下载专区收益统计数据

    Args:
        start_date: 开始日期，格式 YYYY-MM-DD
        end_date: 结束日期，格式 YYYY-MM-DD
        url: 目标网站URL
        download_button_selector: 下载按钮的CSS选择器
        debug_port: Chrome 远程调试端口，用于连接已有的浏览器实例

    Returns:
        下载文件的路径，失败返回None
    """
    start_str = start_date.replace('-', '')
    end_str = end_date.replace('-', '')
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '源数据/export')
    os.makedirs(output_dir, exist_ok=True)

    save_path = os.path.join(output_dir, f'shouyi{start_str}-{end_str}.xlsx')
    if os.path.exists(save_path):
        print(f"收益文件已存在: {save_path}，跳过下载")
        return save_path

    # 根据日期范围计算需要选择的月份
    start_dt = datetime.strptime(start_date, '%Y-%m-%d')
    end_dt = datetime.strptime(end_date, '%Y-%m-%d')
    start_month = start_dt.month
    end_month = end_dt.month
    if start_month <= end_month:
        month_values = list(range(start_month, end_month + 1))
    else:
        month_values = list(range(start_month, 13)) + list(range(1, end_month + 1))
    month_str = ",".join(map(str, month_values))
    print(f"日期范围: {start_date} ~ {end_date}, 选择月份: {month_str}")

    p = None
    browser = None
    try:
        p = sync_playwright().start()

        # 尝试连接到已有的 Chrome 实例
        try:
            print(f"尝试连接到已有 Chrome 实例 (端口: {debug_port})...")
            browser = p.chromium.connect_over_cdp(f"http://localhost:{debug_port}")
            print("成功连接到已有 Chrome 实例")
            # 获取已有的上下文和页面，或创建新页面
            contexts = browser.contexts
            if contexts:
                context = contexts[0]
                page = context.new_page()
            else:
                context = browser.new_context(accept_downloads=True)
                page = context.new_page()
        except Exception as connect_err:
            print(f"连接已有实例失败: {connect_err}，启动新浏览器...")
            browser = p.chromium.launch(headless=False)
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()

        print(f"正在打开: {url}")
        page.goto(url, wait_until="networkidle")
        print("等待页面加载...")
        page.wait_for_timeout(5000)

        # 设置页面缩放为75%
        page.evaluate("document.body.style.zoom = '0.75'")
        page.wait_for_timeout(1000)

        # 检测登录页并截图二维码
        qr_screenshot_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'qr_code_shouyi.png')
        current_url = page.url
        if 'login' in current_url.lower() or 'sso' in current_url.lower() or 'cas' in current_url.lower():
            print("检测到登录页面，点击生成二维码...")
            try:
                page.click("#code", timeout=15000)
                print("已点击 #code，等待二维码加载...")
                page.wait_for_timeout(3000)
                page.screenshot(path=qr_screenshot_path, full_page=False)
                print(f"QR_SCREENSHOT:{qr_screenshot_path}")
                feishu_send_image(qr_screenshot_path, "专区收益统计 登录二维码已生成，请扫码登录：", GROUP_CHAT_ID)
            except Exception as e:
                print(f"点击 #code 失败: {e}，尝试直接截图...")
                page.screenshot(path=qr_screenshot_path, full_page=False)
                print(f"QR_SCREENSHOT:{qr_screenshot_path}")
                feishu_send_image(qr_screenshot_path, "专区收益统计 登录二维码已生成，请扫码登录：", GROUP_CHAT_ID)

            print("请扫描二维码登录，等待中（最多5分钟）...")
            # 等待登录完成（URL变化或菜单出现）
            login_success = False
            for _ in range(300):
                page.wait_for_timeout(1000)
                try:
                    new_url = page.url
                    if 'login' not in new_url.lower() and 'sso' not in new_url.lower() and 'cas' not in new_url.lower():
                        print("检测到页面跳转，已离开登录页")
                        page.wait_for_timeout(5000)
                        login_success = True
                        break
                    if page.query_selector('li[data-id="00050007"]'):
                        print("检测到菜单元素，登录成功！")
                        login_success = True
                        break
                except Exception:
                    pass

            if not login_success:
                print("等待登录超时，请检查是否已扫码登录")
                page.wait_for_timeout(3000)
            else:
                print("登录完成，继续执行...")
                page.wait_for_timeout(2000)
        else:
            print("未检测到登录页，已登录状态，继续执行...")

        # 设置月份
        try:
            page.evaluate(f"mini.get('month').setValue('{month_str}')")
            print(f"设置month值: {month_str}")
            page.wait_for_timeout(1000)
        except Exception as e:
            print(f"设置month值失败: {str(e)}")

        # 点击下载按钮
        print("查找下载按钮...")
        page.wait_for_selector(download_button_selector, timeout=10000)
        print("点击下载按钮...")
        page.click(download_button_selector)
        page.wait_for_timeout(5000)

        # 点击导出按钮
        export_button_selector = "#mini-4-action"
        print("查找导出按钮...")
        page.wait_for_selector(export_button_selector, timeout=10000)
        print("点击导出按钮...")
        with page.expect_download(timeout=120000) as download_info:
            page.click(export_button_selector)

        download = download_info.value
        file_extension = os.path.splitext(download.suggested_filename)[1]
        custom_filename = f"shouyi{start_str}-{end_str}{file_extension}"
        save_path = os.path.join(output_dir, custom_filename)
        download.save_as(save_path)
        print(f"文件已保存到: {save_path}")

        return save_path

    except Exception as e:
        print(f"下载过程中出错: {str(e)}")
        return None


def main(debug_port=9222):
    config = load_config()
    start_date = config['date_range']['start_date']
    end_date = config['date_range']['end_date']

    print(f"\n{'='*60}")
    print(f"收益数据导出")
    print(f"日期范围: {start_date} ~ {end_date}")
    print(f"{'='*60}")

    result = download_revenue(start_date, end_date, debug_port=debug_port)
    if result is None:
        print("收益数据下载失败")
        return False

    print("\n收益数据导出完成")
    return True


if __name__ == "__main__":
    main()
