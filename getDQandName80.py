from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from datetime import datetime
import time
import os
import glob


def main():
    # 配置 Chrome 选项
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    # 设置下载目录
    download_dir = r"D:\MonthReport\中间数据"
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 30)

    try:
        # 1. 打开登录页面
        url = "https://epoint.etrading.cn/qyunitinfo/szcg/login4webnew/login/login"
        print(f"正在打开: {url}")
        driver.get(url)
        time.sleep(2)

        # 2. 输入账号密码并登录
        print("正在输入账号密码...")
        username_input = wait.until(EC.presence_of_element_located((By.ID, "txtUserName$text")))
        username_input.clear()
        username_input.send_keys("yangt")

        password_input = wait.until(EC.presence_of_element_located((By.ID, "txtPwd$text")))
        password_input.clear()
        password_input.send_keys("Dzjy@123")

        # 点击登录按钮
        login_btn = wait.until(EC.element_to_be_clickable((By.ID, "account-login-btn")))
        login_btn.click()

        # 3. 等待登录完成（约15秒）
        print("正在等待登录完成...")
        time.sleep(15)
        print("登录完成，当前URL:", driver.current_url)

        # # 4. 点击"工作台"按钮
        # print("正在点击工作台...")
        # workbench_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".public-btn.zone-workbench")))
        # workbench_btn.click()
        # time.sleep(3)
        # print("已进入工作台")

        # 5. 关闭小广告/引导提示
        print("正在关闭小广告...")
        try:
            skip_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".introjs-skipbutton"))
            )
            skip_btn.click()
            time.sleep(1)
            print("小广告已关闭")
        except:
            print("未发现小广告，继续执行")

        # 6. 点击"配置中心"菜单
        print("正在点击配置中心...")
        config_menu = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-code="1019"] > .menu-link')))
        config_menu.click()
        time.sleep(1)

        # 6. 点击"运营统计"子菜单
        print("正在点击运营统计...")
        stat_menu = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-code="10190017"] > .menu-link')))
        stat_menu.click()
        time.sleep(3)
        print("已进入运营统计页面")

        # 7. 切换到 iframe 并点击"更多"按钮
        print("正在切换到 iframe...")
        # 等待 iframe 加载完成
        time.sleep(3)

        # 获取所有 iframe
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        print(f"找到 {len(iframes)} 个 iframe")

        # 尝试切换到正确的 iframe
        switched = False
        for i, iframe in enumerate(iframes):
            try:
                driver.switch_to.frame(iframe)
                # 检查是否能找到目标元素
                test_elem = driver.find_elements(By.CSS_SELECTOR, '.com-item-hd .hd-condition .hd-more')
                if test_elem:
                    print(f"已切换到第 {i+1} 个 iframe，找到目标元素")
                    switched = True
                    break
                else:
                    driver.switch_to.default_content()
            except:
                driver.switch_to.default_content()
                continue

        if not switched:
            print("未找到正确的 iframe，尝试直接切换第一个 iframe")
            driver.switch_to.frame(0)

        # 点击"更多"按钮（通过父级精确定位）
        print("正在点击更多按钮...")
        more_btn = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, '.com-item-hd .hd-condition .hd-more')
        ))
        more_btn.click()
        time.sleep(2)

        # 8. 切换回主页面，处理弹出的对话框
        driver.switch_to.default_content()
        time.sleep(1)

        # 9. 在弹出的对话框中切换到 iframe
        print("正在处理弹出对话框...")
        # 等待对话框中的 iframe 加载
        dialog_iframe = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, '.mini-window iframe[name^="mini-iframe"]')
        ))
        driver.switch_to.frame(dialog_iframe)

        # 10. 点击"导出"按钮
        print("正在点击导出按钮...")
        export_btn = wait.until(EC.element_to_be_clickable((By.ID, "dataexport")))
        export_btn.click()
        time.sleep(2)

        # 11. 点击第二个"导出"按钮
        print("正在点击确认导出按钮...")
        confirm_export_btn = wait.until(EC.element_to_be_clickable((By.ID, "mini-3-action")))
        confirm_export_btn.click()
        time.sleep(5)

        # 12. 切换回主页面
        driver.switch_to.default_content()

        # 13. 等待文件下载完成
        print("正在等待文件下载...")
        time.sleep(10)

        # 14. 重命名下载的文件
        today = datetime.now().strftime("%Y%m%d")
        new_filename = f"80DQ{today}"

        # 查找最新下载的文件
        files = glob.glob(os.path.join(download_dir, "*"))
        if files:
            latest_file = max(files, key=os.path.getctime)
            file_ext = os.path.splitext(latest_file)[1]
            new_filepath = os.path.join(download_dir, new_filename + file_ext)

            # 如果目标文件已存在，先删除
            if os.path.exists(new_filepath):
                os.remove(new_filepath)

            os.rename(latest_file, new_filepath)
            print(f"文件已重命名为: {new_filepath}")
        else:
            print("未找到下载的文件")

        print("任务完成!")
        print("5秒后自动关闭浏览器...")
        time.sleep(5)

    except Exception as e:
        print(f"发生错误: {e}")
        input("按回车键关闭浏览器...")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
