from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import json

def fetch_video_detail():
    edge_options = Options()
    edge_options.add_argument('--headless')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--disable-dev-shm-usage')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--window-size=1280,900')
    edge_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    driver = webdriver.Edge(options=edge_options)
    
    try:
        url = "https://channels.weixin.qq.com/finder-preview/pages/sph?id=A3WSeH2uBa"
        driver.get(url)
        
        # 等待更长时间让页面完全加载
        time.sleep(8)
        
        # 尝试获取所有文本内容
        body_text = driver.find_element(By.TAG_NAME, 'body').text
        print("=== 页面完整文本 ===")
        print(body_text)
        print("\n" + "="*50)
        
        # 获取完整HTML源码
        html = driver.page_source
        print("\n=== 页面完整HTML（关键部分）===")
        # 查找包含视频信息的script标签
        import re
        scripts = re.findall(r'<script[^>]*>(.*?)</script>', html, re.DOTALL)
        for i, script in enumerate(scripts):
            if 'feed' in script.lower() or 'video' in script.lower() or 'title' in script.lower() or 'desc' in script.lower():
                if len(script) > 50:
                    print(f"\n--- Script {i} ---")
                    print(script[:2000])
        
        # 查找所有元素
        print("\n=== 所有可见文本元素 ===")
        elements = driver.find_elements(By.XPATH, "//*[text()]")
        for elem in elements:
            text = elem.text.strip()
            if text and len(text) > 2:
                tag = elem.tag_name
                classes = elem.get_attribute('class') or ''
                print(f"<{tag} class='{classes}'> {text[:200]}")
        
        # 截图
        driver.save_screenshot("video_detail.png")
        print("\n截图已保存")
        
        # 尝试获取网络请求中的视频信息
        logs = driver.get_log('performance')
        print("\n=== 网络请求中的关键信息 ===")
        for log in logs:
            msg = log.get('message', '')
            if 'finder' in msg or 'feed' in msg or 'video' in msg or 'sph' in msg:
                if 'url' in msg:
                    print(msg[:500])
                    print("---")
        
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()
    finally:
        driver.quit()

if __name__ == "__main__":
    fetch_video_detail()
