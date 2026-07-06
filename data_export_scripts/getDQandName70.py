"""
getDQandName70.py
爬取 etrading.cn 专区列表（共 313 条），通过 Selenium 渲染页面后
直接解析 H5 DOM 元素（tr.mini-grid-row / td），
按分页点击 mini-pager-nextbutton 翻页，结果保存为 Excel。

依赖：
    pip install selenium pandas openpyxl
"""

import sys
import os
import time
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ─── 配置 ───────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
OUTPUT_FILE = os.path.join(PROJECT_DIR, "中间数据", "70DQ.xlsx")
TARGET_URL  = "https://www.etrading.cn/PSPFrame/workbenchmis/pages/regionapply/QuickReference_List"

# 列结构（0-based 在 tr 的所有 td 中）：
#   td[0]=空占位, td[1]=checkbox, td[2]=序号,
#   td[3]=专区编号, td[4]=专区名称, td[5]=所属地区,
#   td[6]=分公司, td[7]=远程交付, td[8]=标证通接入状态,
#   td[9]=标证通(图标), td[10]=详情(图标)
COL_NAMES  = ["专区编号", "专区名称", "所属地区", "分公司", "远程交付", "标证通接入状态"]
COL_INDICES = [3, 4, 5, 6, 7, 8]   # 对应上面列名的 td 索引
# ────────────────────────────────────────────────────────


# ── chromedriver 本地管理 ────────────────────────────────
def _find_local_chromedriver() -> str | None:
    """在 Selenium 缓存目录中查找本地已有的 chromedriver"""
    cache_base = os.path.join(os.environ.get("LOCALAPPDATA", ""),
                              ".cache", "selenium", "chromedriver", "win64")
    if not os.path.exists(cache_base):
        return None
    # 遍历缓存目录，找最新版本的 chromedriver
    for v in sorted(os.listdir(cache_base), reverse=True):
        p = os.path.join(cache_base, v, "chromedriver-win64", "chromedriver.exe")
        if os.path.exists(p):
            print(f"  使用本地缓存 chromedriver: {v}")
            return p
    return None


def create_driver() -> webdriver.Chrome:
    opts = Options()
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)

    driver_path = _find_local_chromedriver()
    if driver_path:
        service = Service(executable_path=driver_path)
        return webdriver.Chrome(service=service, options=opts)

    # 本地无缓存，使用 Selenium Manager 自动管理（仅首次需要网络）
    print("  本地未找到 chromedriver，使用 Selenium Manager 自动配置...")
    return webdriver.Chrome(options=opts)


# ── DOM 解析 ─────────────────────────────────────────────
def wait_for_rows(driver: webdriver.Chrome, timeout: int = 30) -> None:
    """等待表格行渲染完毕"""
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "tr.mini-grid-row"))
    )
    time.sleep(1)


def wait_rows_stable(driver: webdriver.Chrome, timeout: int = 30, stable_sec: float = 2.0) -> None:
    """等待行数稳定（连续 stable_sec 秒内不再增加），确保大页全部渲染完"""
    deadline = time.time() + timeout
    prev_count = -1
    stable_since = None
    while time.time() < deadline:
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.mini-grid-row")
        count = len(rows)
        if count != prev_count:
            prev_count = count
            stable_since = time.time()
        elif stable_since and time.time() - stable_since >= stable_sec:
            return   # 行数稳定，加载完成
        time.sleep(0.5)


def parse_current_page(driver: webdriver.Chrome, existing_map: dict) -> int:
    """解析当前页所有渲染出来的行，并存入 existing_map。返回新增数量。"""
    rows = driver.find_elements(By.CSS_SELECTOR, "tr.mini-grid-row")
    added = 0
    for row in rows:
        tds = row.find_elements(By.TAG_NAME, "td")
        if len(tds) <= max(COL_INDICES):
            continue
        record = {}
        for col_name, td_idx in zip(COL_NAMES, COL_INDICES):
            td = tds[td_idx]
            inner = td.find_elements(By.CSS_SELECTOR, ".mini-grid-cell-inner")
            text = inner[0].text.strip() if inner else td.text.strip()
            record[col_name] = text if text and text != "\xa0" else ""
        
        # 使用"专区编号"作为唯一键
        key = record.get("专区编号")
        if key and key not in existing_map:
            existing_map[key] = record
            added += 1
    return added


def has_next_page(driver: webdriver.Chrome) -> bool:
    """判断下一页按钮是否可点击（不含 disabled）"""
    btns = driver.find_elements(By.CSS_SELECTOR, "a.mini-pager-nextbutton")
    for btn in btns:
        cls = btn.get_attribute("class") or ""
        if "mini-button-disabled" not in cls:
            return True
    return False


def click_next_page(driver: webdriver.Chrome) -> None:
    """点击下一页"""
    btns = driver.find_elements(By.CSS_SELECTOR, "a.mini-pager-nextbutton")
    for btn in btns:
        cls = btn.get_attribute("class") or ""
        if "mini-button-disabled" not in cls:
            try:
                btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn)
            return


def set_page_size_via_ui(driver: webdriver.Chrome, size: int) -> None:
    """点击 UI 修改每页条数"""
    print(f"  尝试点击 UI 设置每页条数为 {size}...")
    try:
        # 等待分页大小选择框出现
        combobox = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".mini-pager-size .mini-buttonedit"))
        )
        trigger = combobox.find_element(By.CSS_SELECTOR, ".mini-buttonedit-button")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", trigger)
        trigger.click()
        time.sleep(1)

        # 在弹出的 listbox 中找到对应选项
        options = driver.find_elements(By.CSS_SELECTOR, ".mini-listbox-item")
        for opt in options:
            if str(size) == opt.text.strip():
                opt.click()
                print(f"  已点击选项 {size}")
                return
        print(f"  未能直接找到选项 {size}，尝试通过 JS 补救...")
        driver.execute_script(f"var g=mini.get('datagrid'); if(g) g.setPageSize({size});")
    except Exception as e:
        print(f"  UI 交互失败: {e}，尝试通过 JS 补救...")
        driver.execute_script(f"var g=mini.get('datagrid'); if(g) g.setPageSize({size});")


def scroll_to_load_all(driver: webdriver.Chrome) -> None:
    """滚动表格容器以触发布局加载（针对虚拟滚动/懒加载）"""
    print("  正在滚动加载数据...")
    try:
        container = driver.find_element(By.CSS_SELECTOR, ".mini-grid-rows-view")
        last_top = -1
        while True:
            # 每次向下滚动 600 像素
            driver.execute_script("arguments[0].scrollTop += 600;", container)
            time.sleep(1.0)
            curr_top = driver.execute_script("return arguments[0].scrollTop;", container)
            if curr_top == last_top: # 到底了
                break
            last_top = curr_top
    except Exception as e:
        print(f"  滚动失败: {e}")


def get_pager_info(driver: webdriver.Chrome) -> str:
    """获取分页信息文字，例如\"每页 270 条, 共 313 条\""""
    try:
        el = driver.find_element(By.CSS_SELECTOR, ".mini-pager-right")
        return el.text.strip()
    except Exception:
        return ""


# ── 主流程 ───────────────────────────────────────────────
def crawl() -> list[dict]:
    driver = create_driver()
    all_records: list[dict] = []

    try:
        print(f"打开页面: {TARGET_URL}")
        driver.get(TARGET_URL)

        print("等待表格加载...")
        wait_for_rows(driver, timeout=30)

        # 1. 设置页码为 270
        set_page_size_via_ui(driver, 150)
        time.sleep(2)
        wait_rows_stable(driver, timeout=30, stable_sec=2.0)
        print(f"分页信息: {get_pager_info(driver)}")

        page = 1
        all_records_map = {} # 专区编号 -> record
        while True:
            print(f"  正在处理第 {page} 页...")
            
            # 由于是虚拟滚动/懒加载，需要边滚动边收集
            container = driver.find_element(By.CSS_SELECTOR, ".mini-grid-rows-view")
            last_top = -1
            added_this_page = 0
            while True:
                # 解析当前可见行
                new_count = parse_current_page(driver, all_records_map)
                added_this_page += new_count
                
                # 滚动
                driver.execute_script("arguments[0].scrollTop += 800;", container)
                time.sleep(1.0)
                curr_top = driver.execute_script("return arguments[0].scrollTop;", container)
                if curr_top == last_top:
                    # 到底了，最后再整页解析一次确保没漏
                    parse_current_page(driver, all_records_map)
                    break
                last_top = curr_top
            
            print(f"    第 {page} 页处理完毕，累计收集 {len(all_records_map)} 条记录")

            if not has_next_page(driver):
                print("  已到最后一页")
                break

            click_next_page(driver)
            page += 1
            time.sleep(2)
            wait_rows_stable(driver, timeout=30, stable_sec=2.0)

        all_records = list(all_records_map.values())
        print(f"\n✅ 共获取 {len(all_records)} 条记录")
    finally:
        driver.quit()

    return all_records


def save_to_excel(records: list[dict]) -> None:
    if not records:
        print("⚠️  没有数据可保存")
        return
    df = pd.DataFrame(records, columns=COL_NAMES)
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"✅ 已保存 {len(df)} 行数据到: {OUTPUT_FILE}")


def main():
    records = crawl()
    save_to_excel(records)


if __name__ == "__main__":
    main()
