from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import json

def fetch_video_content():
    edge_options = Options()
    edge_options.add_argument('--headless')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--disable-dev-shm-usage')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--window-size=1280,800')
    
    driver = webdriver.Edge(options=edge_options)
    
    try:
        url = "https://channels.weixin.qq.com/finder-preview/pages/sph?id=A3WSeH2uBa"
        driver.get(url)
        
        # 等待页面加载
        time.sleep(5)
        
        # 获取页面源码
        html = driver.page_source
        print("=== 页面HTML ===")
        print(html[:5000])
        
        # 尝试找到视频标题
        try:
            title = driver.find_element(By.CSS_SELECTOR, '.finder-preview-title, .video-title, h1, .title').text
            print(f"\n=== 视频标题 ===\n{title}")
        except:
            print("\n未找到标题")
        
        # 尝试找到视频描述
        try:
            desc = driver.find_element(By.CSS_SELECTOR, '.finder-preview-desc, .video-desc, .description').text
            print(f"\n=== 视频描述 ===\n{desc}")
        except:
            print("\n未找到描述")
        
        # 截图保存
        driver.save_screenshot("video_screenshot.png")
        print("\n截图已保存到 video_screenshot.png")
        
        # 尝试获取所有文本内容
        body_text = driver.find_element(By.TAG_NAME, 'body').text
        print(f"\n=== 页面文本内容 ===\n{body_text[:3000]}")
        
    except Exception as e:
        print(f"错误: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    fetch_video_content()
