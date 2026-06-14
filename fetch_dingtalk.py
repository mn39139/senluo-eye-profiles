from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def fetch_dingtalk_content():
    edge_options = Options()
    edge_options.add_argument('--headless')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--disable-dev-shm-usage')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--window-size=1280,900')
    edge_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    driver = webdriver.Edge(options=edge_options)
    
    try:
        url = "https://n.dingtalk.com/dingding/dingding-marketing/ai-marketing/index.html#/share/case/24581"
        driver.get(url)
        
        time.sleep(8)
        
        body_text = driver.find_element(By.TAG_NAME, 'body').text
        print("=== 页面完整文本 ===")
        print(body_text[:5000])
        print("\n" + "="*50)
        
        driver.save_screenshot("dingtalk_screenshot.png")
        print("截图已保存")
        
        html = driver.page_source
        print("\n=== HTML关键内容 ===")
        print(html[:3000])
        
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()
    finally:
        driver.quit()

if __name__ == "__main__":
    fetch_dingtalk_content()
