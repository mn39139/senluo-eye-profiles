from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from PIL import Image
import os
import time

def html_to_jpg():
    html_path = os.path.abspath("D:\\TRAE\\haibao\\new_poster.html")
    output_path = "D:\\TRAE\\haibao\\poster.jpg"
    
    # 设置Edge选项
    edge_options = Options()
    edge_options.add_argument('--headless')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--disable-dev-shm-usage')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--window-size=750,1334')
    
    # 尝试多个可能的Edge路径
    edge_paths = [
        "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
        "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
        os.path.expanduser("~\\AppData\\Local\\Microsoft\\Edge\\Application\\msedge.exe"),
    ]
    
    edge_binary = None
    for path in edge_paths:
        if os.path.exists(path):
            edge_binary = path
            break
    
    if edge_binary:
        edge_options.binary_location = edge_binary
    
    # 尝试使用系统EdgeDriver
    edgedriver_paths = [
        "C:\\Windows\\msedgedriver.exe",
        "C:\\Windows\\System32\\msedgedriver.exe",
        os.path.expanduser("~\\msedgedriver.exe"),
        "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedgedriver.exe",
        "C:\\Program Files\\Microsoft\\Edge\\Application\\msedgedriver.exe",
    ]
    
    driver = None
    for ed_path in edgedriver_paths:
        if os.path.exists(ed_path):
            service = Service(ed_path)
            driver = webdriver.Edge(service=service, options=edge_options)
            break
    
    if driver is None:
        # 尝试默认方式
        driver = webdriver.Edge(options=edge_options)
    
    try:
        # 加载HTML文件
        driver.get(f"file:///{html_path}")
        
        # 等待页面加载和字体渲染
        time.sleep(3)
        
        # 获取海报元素
        poster = driver.find_element('css selector', '.poster')
        
        # 截图
        poster.screenshot("temp.png")
        
        # 转换为JPG
        img = Image.open("temp.png")
        img = img.convert('RGB')
        img.save(output_path, 'JPEG', quality=95)
        
        # 清理临时文件
        os.remove("temp.png")
        
        print(f"JPG已保存到: {output_path}")
        
    finally:
        driver.quit()

if __name__ == "__main__":
    html_to_jpg()
