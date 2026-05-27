import os
import time
import json
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

# ========== 配置区域 ==========
search_keyword = "失信人"
driver_path = r"D:\Files\EdgDownload\edgedriver_win64\msedgedriver.exe"
save_dir = r'D:\This computer\desktop\失信人文件'
MAX_DOWNLOADS = 30  # 下载30个PDF后停止

os.makedirs(save_dir, exist_ok=True)

# ========== 初始化Edge浏览器 ==========
print("=" * 60)
print("PDF自动下载工具 - 修正版")
print(f"搜索关键词: '{search_keyword}'")
print(f"目标下载数量: {MAX_DOWNLOADS}个")
print(f"保存路径: {save_dir}")
print("=" * 60)

# 创建Edge选项
options = Options()

# 关键设置：让PDF在浏览器中打开，而不是直接下载
prefs = {
    "download.default_directory": save_dir,
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": False,  # False = 在浏览器中打开PDF
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

# 显示浏览器窗口，便于调试
options.add_argument('--start-maximized')

# 禁用自动化特征
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

try:
    print("\n正在初始化Edge浏览器...")
    
    if not os.path.exists(driver_path):
        print(f"✗ Edge驱动文件不存在: {driver_path}")
        exit()
    
    service = Service(executable_path=driver_path)
    driver = webdriver.Edge(service=service, options=options)
    
    driver.set_page_load_timeout(45)
    driver.set_script_timeout(45)
    
    print(f"✓ Edge浏览器初始化成功！")
    
except Exception as e:
    print(f"\n✗ 浏览器初始化失败: {e}")
    exit()

# 创建等待对象
wait = WebDriverWait(driver, 20)

# ========== 访问网站并等待手动搜索 ==========
print(f"\n正在访问网站...")
driver.get("https://www.neeq.com.cn/index/searchInfo.do")
print("✓ 网站已加载")

print("\n" + "="*60)
print("请手动完成以下操作:")
print("1. 在搜索框中输入搜索条件")
print("2. 点击搜索按钮")
print("3. 等待搜索结果加载完成")
print("="*60 + "\n")

input("完成后请按回车键继续...")

# ========== 创建URL记录文件 ==========
url_record_file = os.path.join(save_dir, f"pdf_urls_{time.strftime('%Y%m%d_%H%M%S')}.json")
url_records = []
print(f"\nPDF URL将保存到: {url_record_file}")

# ========== 主要下载逻辑 ==========
print(f"\n开始下载PDF文件，目标: {MAX_DOWNLOADS}个")

total_downloaded = 0
page_num = 1
failed_downloads = []

while total_downloaded < MAX_DOWNLOADS:
    print(f"\n=== 正在处理第 {page_num} 页 ===")
    
    # 等待当前页的搜索结果加载
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//p[@class='tit1']")))
        time.sleep(2)  # 额外等待确保完全加载
    except TimeoutException:
        print("✗ 等待搜索结果超时")
        break
    
    # 获取当前页所有标题元素
    title_elements = driver.find_elements(By.XPATH, "//p[@class='tit1']")
    print(f"本页找到 {len(title_elements)} 个结果")
    
    if len(title_elements) == 0:
        print("当前页面没有结果")
        break
    
    # 获取主窗口句柄
    main_window = driver.current_window_handle
    
    # 遍历当前页所有标题
    for idx, title_element in enumerate(title_elements, 1):
        # 检查是否已达到目标数量
        if total_downloaded >= MAX_DOWNLOADS:
            print(f"已达到目标数量 {MAX_DOWNLOADS} 个，停止处理")
            break
        
        try:
            # 获取标题文本
            full_title = title_element.get_attribute("title")
            if not full_title:
                full_title = title_element.text
            
            # 清理文件名
            safe_filename = "".join(c for c in full_title if c not in r'<>:"/\\|?*')
            if len(safe_filename) > 80:
                safe_filename = safe_filename[:80]
            
            print(f"\n处理第 {idx} 个: {safe_filename[:60]}...")
            print(f"当前已下载: {total_downloaded}/{MAX_DOWNLOADS}")
            
            # 记录点击前的窗口
            windows_before = driver.window_handles
            
            # 点击标题
            print("  正在点击标题...")
            title_element.click()
            time.sleep(3)  # 等待新窗口打开
            
            # 获取所有窗口句柄
            windows_after = driver.window_handles
            
            # 查找新打开的窗口
            new_window = None
            for window in windows_after:
                if window not in windows_before:
                    new_window = window
                    break
            
            if new_window:
                # 切换到新窗口
                driver.switch_to.window(new_window)
                
                # 获取PDF浏览界面的URL（关键步骤）
                pdf_url = driver.current_url
                print(f"  PDF浏览界面URL: {pdf_url}")
                
                # 记录URL信息
                url_record = {
                    "index": total_downloaded + 1,
                    "title": full_title,
                    "safe_filename": safe_filename,
                    "pdf_url": pdf_url,
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "is_pdf": pdf_url.lower().endswith('.pdf')
                }
                
                # 尝试下载PDF
                if pdf_url.lower().endswith('.pdf'):
                    print(f"  ✓ 这是PDF直接链接")
                    
                    try:
                        # 设置请求头
                        headers = {
                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36 Edg/91.0.864.59',
                            'Referer': 'https://www.neeq.com.cn/'
                        }
                        
                        # 获取浏览器cookies
                        selenium_cookies = driver.get_cookies()
                        cookies_dict = {}
                        for cookie in selenium_cookies:
                            cookies_dict[cookie['name']] = cookie['value']
                        
                        # 下载PDF
                        response = requests.get(pdf_url, headers=headers, cookies=cookies_dict, timeout=30)
                        
                        if response.status_code == 200:
                            # 创建文件名（添加序号防止重复）
                            pdf_filename = f"{total_downloaded+1:03d}_{safe_filename}.pdf"
                            pdf_path = os.path.join(save_dir, pdf_filename)
                            
                            # 保存文件
                            with open(pdf_path, 'wb') as f:
                                f.write(response.content)
                            
                            print(f"  ✓ 已下载: {pdf_filename}")
                            total_downloaded += 1
                            url_record["download_status"] = "success"
                            url_record["saved_filename"] = pdf_filename
                        else:
                            print(f"  ✗ 下载失败，状态码: {response.status_code}")
                            url_record["download_status"] = "failed"
                            failed_downloads.append(safe_filename)
                            
                    except Exception as e:
                        print(f"  ✗ 下载失败: {e}")
                        url_record["download_status"] = "failed"
                        failed_downloads.append(safe_filename)
                else:
                    print(f"  ⚠️ 不是PDF链接，跳过")
                    url_record["download_status"] = "not_pdf"
                    failed_downloads.append(safe_filename)
                
                # 保存URL记录
                url_records.append(url_record)
                
                # 关闭当前窗口，返回主窗口
                driver.close()
                driver.switch_to.window(main_window)
                print("  已返回主列表页面")
                
            else:
                print("  ⚠️ 未打开新窗口，PDF可能已在当前窗口打开")
                # 检查当前URL是否为PDF
                current_url = driver.current_url
                if current_url.lower().endswith('.pdf'):
                    print(f"  ✓ 当前窗口是PDF链接: {current_url}")
                    # 可以尝试下载，但需要返回搜索结果页面
                    # 这里为了简化，先跳过
                # 返回搜索结果页面
                driver.get("https://www.neeq.com.cn/index/searchInfo.do")
                print("  已返回搜索结果页面")
                time.sleep(3)
                
        except Exception as e:
            print(f"  处理出错: {e}")
            failed_downloads.append(full_title if 'full_title' in locals() else f"结果{idx}")
        
        # 短暂等待，避免操作过快
        time.sleep(2)
    
    # 保存当前的URL记录
    try:
        with open(url_record_file, 'w', encoding='utf-8') as f:
            json.dump(url_records, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存URL记录失败: {e}")
    
    # 检查是否已达到目标数量
    if total_downloaded >= MAX_DOWNLOADS:
        break
    
    # ========== 翻页处理 ==========
    print(f"\n当前页处理完成，尝试翻页...")
    
    try:
        # 查找"下一页"按钮
        next_button = driver.find_element(By.XPATH, "//a[@class='next']")
        
        # 检查按钮是否可用
        if "disabled" in next_button.get_attribute("class"):
            print("已到达最后一页，无法继续翻页")
            break
        
        # 点击下一页
        print("点击下一页按钮...")
        next_button.click()
        page_num += 1
        
        # 等待新页面加载
        print("等待新页面加载...")
        time.sleep(5)
        
        # 等待新页面的结果出现
        wait.until(EC.presence_of_element_located((By.XPATH, "//p[@class='tit1']")))
        print(f"第 {page_num} 页加载成功")
        
    except NoSuchElementException:
        print("未找到'下一页'按钮，可能已到最后一页")
        break
    except Exception as e:
        print(f"翻页时出错: {e}")
        break

# ========== 最终统计和保存 ==========
print(f"\n{'='*60}")
print("任务完成!")
print(f"{'='*60}")
print(f"统计信息:")
print(f"- 目标下载数量: {MAX_DOWNLOADS}")
print(f"- 实际下载数量: {total_downloaded}")
print(f"- 失败数量: {len(failed_downloads)}")
print(f"- 处理页面数: {page_num}")

if failed_downloads:
    print(f"\n失败的文件列表:")
    for i, filename in enumerate(failed_downloads[:10], 1):
        print(f"  {i}. {filename[:50]}...")

# 保存最终的URL记录
print(f"\n正在保存PDF URL记录...")
try:
    with open(url_record_file, 'w', encoding='utf-8') as f:
        json.dump(url_records, f, ensure_ascii=False, indent=2)
    print(f"✓ PDF URL记录已保存到: {url_record_file}")
    
    # 同时保存为文本文件便于查看
    txt_file = url_record_file.replace('.json', '.txt')
    with open(txt_file, 'w', encoding='utf-8') as f:
        f.write("PDF URL记录\n")
        f.write("="*50 + "\n")
        f.write(f"记录时间: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"搜索关键词: {search_keyword}\n")
        f.write(f"下载数量: {total_downloaded}/{MAX_DOWNLOADS}\n")
        f.write("="*50 + "\n\n")
        
        for record in url_records:
            f.write(f"[{record.get('index', 'N/A')}] {record.get('title', '')[:50]}...\n")
            f.write(f"  URL: {record.get('pdf_url', 'N/A')}\n")
            f.write(f"  状态: {record.get('download_status', 'unknown')}\n")
            f.write(f"  文件名: {record.get('saved_filename', 'N/A')}\n")
            f.write("-"*50 + "\n")
    
    print(f"✓ 文本格式记录已保存到: {txt_file}")
    
except Exception as e:
    print(f"✗ 保存记录失败: {e}")

print(f"\n所有PDF文件保存到: {save_dir}")
print("建议检查下载的文件是否完整")

# 询问是否打开保存目录
response = input("\n是否要打开保存目录？(y/n): ")
if response.lower() == 'y':
    try:
        os.startfile(save_dir)
        print(f"已打开目录: {save_dir}")
    except:
        print("无法自动打开目录，请手动访问")

input("\n按回车键关闭浏览器...")
driver.quit()
