"""通过Chrome远程调试协议操作B站——搜索影视飓风，进第一个视频点赞"""
import subprocess, time, os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
DEBUG_PORT = 9222

# 步骤1：以调试模式启动 Chrome（使用用户自己的配置，保留登录状态）
print("以调试模式启动 Chrome...")
subprocess.Popen([
    CHROME_PATH,
    f"--remote-debugging-port={DEBUG_PORT}",
    "--no-first-run",
    "--no-default-browser-check",
])
time.sleep(3)

# 步骤2：Selenium 连接到已有 Chrome 实例
print("连接到 Chrome...")
options = Options()
options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")

driver_path = r"C:\Users\33164\Desktop\todo\chromedriver\chromedriver-win64\chromedriver.exe"
service = Service(executable_path=driver_path)

driver = webdriver.Chrome(service=service, options=options)
print(f"已连接，当前页面: {driver.current_url}")
wait = WebDriverWait(driver, 10)

# 步骤3：打开搜索页
print("打开B站搜索影视飓风...")
driver.get("https://search.bilibili.com/all?keyword=%E5%BD%B1%E8%A7%86%E9%A3%93%E9%A3%8E&order=click")
time.sleep(3)

# 步骤4：点击第一个视频
print("查找第一个视频...")
try:
    # B站搜索结果有多种页面结构
    selectors = [
        ".bili-video-card a",
        ".video-list-item a",
        ".search-video-card a",
        ".bili-video-card__wrap a",
    ]
    first_video = None
    for sel in selectors:
        try:
            first_video = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
            break
        except:
            continue

    if not first_video:
        first_video = driver.find_element(By.CSS_SELECTOR, "a[href*='video']")

    href = first_video.get_attribute("href")
    print(f"第一个视频: {href}")
    first_video.click()
    time.sleep(4)
except Exception as e:
    print(f"点击失败: {e}")
    # JS兜底
    driver.execute_script(
        "var a = document.querySelector('.bili-video-card a, .video-list-item a, a[href*=\"video\"]');"
        "if(a) a.click();"
    )
    time.sleep(4)

# 步骤5：切换到视频页（可能在新标签页）
if len(driver.window_handles) > 1:
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(2)
elif "/video/" not in driver.current_url:
    time.sleep(2)

print(f"当前页面: {driver.current_url}")
print(f"页面标题: {driver.title}")

# 步骤6：点赞
print("查找点赞按钮...")
like_selectors = [
    ".video-like .video-like-info",
    ".video-toolbar-left .video-like",
    'span.video-like-info',
    '.video-info-container .video-like',
    '[class*="video-like"]',
]

try:
    like_btn = None
    for sel in like_selectors:
        try:
            candidates = driver.find_elements(By.CSS_SELECTOR, sel)
            for btn in candidates:
                if btn.is_displayed() and btn.size["width"] > 0:
                    like_btn = btn
                    break
            if like_btn:
                break
        except:
            continue

    if like_btn is None:
        # JS查找
        like_btn = driver.execute_script(
            "return document.querySelector('.video-like-info, [class*=\"video-like\"]');"
        )
        if like_btn:
            print("通过JS找到点赞按钮")

    if like_btn is None:
        raise Exception("找不到点赞按钮")

    # 检查是否已点赞
    try:
        class_attr = like_btn.get_attribute("class") or ""
        parent = like_btn.find_element(By.XPATH, "..")
        parent_class = parent.get_attribute("class") or ""
        full_class = class_attr + " " + parent_class
    except:
        full_class = like_btn.get_attribute("class") or ""

    if "on" in full_class or "active" in full_class:
        print("已经点过赞了！")
    else:
        like_btn.click()
        print("点赞成功！")
        time.sleep(1)
except Exception as e:
    print(f"点赞出错: {e}")
    # 终极兜底：JS点击
    try:
        driver.execute_script(
            "var b = document.querySelector('.video-like-info');"
            "if(!b) b = document.querySelector('[class*=\"video-like\"]');"
            "if(b) b.click();"
            "else console.log('no like btn');"
        )
        print("JS兜底执行完成")
    except Exception as e2:
        print(f"兜底也失败: {e2}")
        print("请在页面上手动点赞吧~")

time.sleep(2)
print("完成！浏览器保持打开。")
