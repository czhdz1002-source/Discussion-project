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
MAX_DOWNLOADS = 500  # 下载500个PDF后停止
RETRY_LIMIT = 3      # 每个文件最大重试次数
REQUEST_TIMEOUT = 30  # 请求超时时间

os.makedirs(save_dir, exist_ok=True)

# ========== 初始化Edge浏览器 ==========
print("=" * 70)
print("PDF批量下载工具 - 500文件版本")
print(f"搜索关键词: '{search_keyword}'")
print(f"目标下载数量: {MAX_DOWNLOADS}个")
print(f"保存路径: {save_dir}")
print(f"开始时间: {time.strftime('%Y-%m-%d %H:%M:%S')}")
print("=" * 70)

# 创建Edge选项
options = Options()

# 关键设置：让PDF在浏览器中打开
prefs = {
    "download.default_directory": save_dir,
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": False,  # 在浏览器中打开PDF
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_setting_values.automatic_downloads": 1  # 允许自动下载
}
options.add_experimental_option("prefs", prefs)

# 浏览器优化设置
options.add_argument('--start-maximized')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# 可选：启用无头模式（不显示浏览器窗口）
# options.add_argument('--headless')

try:
    print("\n正在初始化Edge浏览器...")
    
    if not os.path.exists(driver_path):
        print(f"✗ Edge驱动文件不存在: {driver_path}")
        exit()
    
    service = Service(executable_path=driver_path)
    driver = webdriver.Edge(service=service, options=options)
    
    # 设置超时时间
    driver.set_page_load_timeout(60)
    driver.set_script_timeout(60)
    
    print(f"✓ Edge浏览器初始化成功！")
    
except Exception as e:
    print(f"\n✗ 浏览器初始化失败: {e}")
    exit()

# 创建等待对象
wait = WebDriverWait(driver, 30)

# ========== 访问网站并等待手动搜索 ==========
print(f"\n正在访问目标网站...")
try:
    driver.get("https://www.neeq.com.cn/index/searchInfo.do")
    print("✓ 网站已加载")
except Exception as e:
    print(f"✗ 网站访问失败: {e}")
    driver.quit()
    exit()

print("\n" + "="*70)
print("请手动完成以下操作:")
print("1. 在搜索框中输入搜索条件")
print("2. 点击搜索按钮")
print("3. 等待搜索结果加载完成")
print("注意：请确保搜索结果页面正常显示")
print("="*70 + "\n")

input("完成后请按回车键继续...")

# ========== 创建记录文件 ==========
timestamp = time.strftime('%Y%m%d_%H%M%S')
url_record_file = os.path.join(save_dir, f"pdf_urls_{timestamp}.json")
download_log_file = os.path.join(save_dir, f"download_log_{timestamp}.txt")
checkpoint_file = os.path.join(save_dir, f"checkpoint_{timestamp}.json")

url_records = []
downloaded_files = []

# 尝试加载检查点（如果程序中断后重新启动）
if os.path.exists(checkpoint_file):
    try:
        with open(checkpoint_file, 'r', encoding='utf-8') as f:
            checkpoint_data = json.load(f)
            total_downloaded = checkpoint_data.get('total_downloaded', 0)
            page_num = checkpoint_data.get('page_num', 1)
            url_records = checkpoint_data.get('url_records', [])
            downloaded_files = checkpoint_data.get('downloaded_files', [])
        
        print(f"✓ 从检查点恢复: 已下载 {total_downloaded} 个文件，当前第 {page_num} 页")
        print(f"恢复后目标: 下载 {MAX_DOWNLOADS - total_downloaded} 个文件")
    except:
        total_downloaded = 0
        page_num = 1
        print("✗ 检查点文件损坏，从头开始")
else:
    total_downloaded = 0
    page_num = 1

print(f"\nPDF URL将保存到: {url_record_file}")
print(f"下载日志将保存到: {download_log_file}")

# 日志函数
def log_message(message, log_type="INFO"):
    timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
    log_entry = f"[{timestamp}] [{log_type}] {message}"
    
    print(log_entry)
    
    # 写入日志文件
    try:
        with open(download_log_file, 'a', encoding='utf-8') as f:
            f.write(log_entry + "\n")
    except:
        pass
    
    return log_entry

# 保存检查点函数
def save_checkpoint():
    checkpoint_data = {
        'total_downloaded': total_downloaded,
        'page_num': page_num,
        'url_records': url_records,
        'downloaded_files': downloaded_files,
        'timestamp': time.strftime('%Y-%m-%d %H:%M:%S')
    }
    
    try:
        with open(checkpoint_file, 'w', encoding='utf-8') as f:
            json.dump(checkpoint_data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        log_message(f"保存检查点失败: {e}", "ERROR")
        return False

# ========== 主要下载逻辑 ==========
log_message(f"开始下载PDF文件，目标: {MAX_DOWNLOADS}个")

failed_downloads = []
consecutive_failures = 0
max_consecutive_failures = 5  # 连续失败5次后暂停

while total_downloaded < MAX_DOWNLOADS:
    log_message(f"=== 正在处理第 {page_num} 页 ===")
    
    # 等待当前页的搜索结果加载
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//p[@class='tit1']")))
        time.sleep(3)  # 额外等待确保完全加载
    except TimeoutException:
        log_message("等待搜索结果超时，尝试刷新页面...", "WARNING")
        
        # 尝试刷新页面
        try:
            driver.refresh()
            time.sleep(5)
            wait.until(EC.presence_of_element_located((By.XPATH, "//p[@class='tit1']")))
            log_message("页面刷新成功，继续处理")
        except:
            log_message("无法加载搜索结果，可能已到最后一页或网络问题", "ERROR")
            break
    
    # 获取当前页所有标题元素
    try:
        title_elements = driver.find_elements(By.XPATH, "//p[@class='tit1']")
        log_message(f"本页找到 {len(title_elements)} 个结果")
    except Exception as e:
        log_message(f"获取标题元素失败: {e}", "ERROR")
        break
    
    if len(title_elements) == 0:
        log_message("当前页面没有结果，可能已到最后一页", "WARNING")
        break
    
    # 获取主窗口句柄
    main_window = driver.current_window_handle
    
    # 遍历当前页所有标题
    for idx, title_element in enumerate(title_elements, 1):
        # 检查是否已达到目标数量
        if total_downloaded >= MAX_DOWNLOADS:
            log_message(f"已达到目标数量 {MAX_DOWNLOADS} 个，停止处理")
            break
        
        # 跳过已下载的文件（从检查点恢复时）
        current_index = total_downloaded + 1
        if current_index <= len(downloaded_files):
            log_message(f"跳过已下载的文件 {current_index}/{MAX_DOWNLOADS}")
            total_downloaded += 1
            continue
        
        retry_count = 0
        download_success = False
        
        while retry_count < RETRY_LIMIT and not download_success:
            retry_count += 1
            if retry_count > 1:
                log_message(f"第 {retry_count} 次重试下载...", "WARNING")
            
            try:
                # 获取标题文本
                full_title = title_element.get_attribute("title")
                if not full_title:
                    full_title = title_element.text
                
                # 清理文件名
                safe_filename = "".join(c for c in full_title if c not in r'<>:"/\\|?*')
                if len(safe_filename) > 80:
                    safe_filename = safe_filename[:80]
                
                log_message(f"处理第 {current_index} 个: {safe_filename[:60]}...")
                log_message(f"当前进度: {total_downloaded}/{MAX_DOWNLOADS}")
                
                # 记录点击前的窗口
                windows_before = driver.window_handles
                
                # 点击标题
                log_message("正在点击标题...")
                title_element.click()
                time.sleep(4)  # 等待新窗口打开
                
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
                    
                    # 获取PDF浏览界面的URL
                    pdf_url = driver.current_url
                    log_message(f"PDF浏览界面URL: {pdf_url}")
                    
                    # 记录URL信息
                    url_record = {
                        "index": current_index,
                        "title": full_title,
                        "safe_filename": safe_filename,
                        "pdf_url": pdf_url,
                        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                        "is_pdf": pdf_url.lower().endswith('.pdf'),
                        "retry_count": retry_count
                    }
                    
                    # 尝试下载PDF
                    if pdf_url.lower().endswith('.pdf'):
                        log_message(f"这是PDF直接链接，开始下载...")
                        
                        try:
                            # 设置请求头
                            headers = {
                                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0',
                                'Referer': 'https://www.neeq.com.cn/',
                                'Accept': 'application/pdf, */*',
                                'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
                            }
                            
                            # 获取浏览器cookies
                            selenium_cookies = driver.get_cookies()
                            cookies_dict = {}
                            for cookie in selenium_cookies:
                                cookies_dict[cookie['name']] = cookie['value']
                            
                            # 下载PDF
                            response = requests.get(pdf_url, headers=headers, cookies=cookies_dict, 
                                                   timeout=REQUEST_TIMEOUT, stream=True)
                            
                            if response.status_code == 200:
                                # 创建文件名（添加序号防止重复）
                                pdf_filename = f"{current_index:04d}_{safe_filename}.pdf"
                                pdf_path = os.path.join(save_dir, pdf_filename)
                                
                                # 保存文件（使用流式写入，支持大文件）
                                with open(pdf_path, 'wb') as f:
                                    for chunk in response.iter_content(chunk_size=8192):
                                        if chunk:
                                            f.write(chunk)
                                
                                file_size = os.path.getsize(pdf_path) / 1024  # KB
                                log_message(f"✓ 已下载: {pdf_filename} ({file_size:.1f} KB)")
                                
                                total_downloaded += 1
                                downloaded_files.append(pdf_filename)
                                url_record["download_status"] = "success"
                                url_record["saved_filename"] = pdf_filename
                                url_record["file_size_kb"] = file_size
                                
                                download_success = True
                                consecutive_failures = 0  # 重置连续失败计数
                            else:
                                log_message(f"下载失败，状态码: {response.status_code}", "ERROR")
                                url_record["download_status"] = f"failed_http_{response.status_code}"
                                
                        except requests.exceptions.Timeout:
                            log_message("请求超时", "ERROR")
                            url_record["download_status"] = "timeout"
                        except Exception as e:
                            log_message(f"下载失败: {e}", "ERROR")
                            url_record["download_status"] = f"failed_{type(e).__name__}"
                    else:
                        log_message(f"⚠️ 不是PDF链接，跳过", "WARNING")
                        url_record["download_status"] = "not_pdf"
                    
                    # 保存URL记录
                    url_records.append(url_record)
                    
                    # 关闭当前窗口，返回主窗口
                    driver.close()
                    driver.switch_to.window(main_window)
                    log_message("已返回主列表页面")
                    
                else:
                    log_message("⚠️ 未打开新窗口，PDF可能已在当前窗口打开", "WARNING")
                    
                    # 检查当前URL是否为PDF
                    current_url = driver.current_url
                    if current_url.lower().endswith('.pdf'):
                        log_message(f"当前窗口是PDF链接: {current_url}")
                        # 可以尝试下载，但需要返回搜索结果页面
                        # 这里为了简化，先跳过
                    
                    # 尝试返回搜索结果页面
                    try:
                        driver.back()
                        time.sleep(3)
                        # 重新获取标题元素
                        title_elements = driver.find_elements(By.XPATH, "//p[@class='tit1']")
                        log_message("已返回搜索结果页面")
                    except:
                        log_message("返回搜索结果页面失败", "ERROR")
                
                # 保存检查点（每隔5个文件保存一次）
                if total_downloaded % 5 == 0:
                    save_checkpoint()
                    log_message(f"检查点已保存，当前进度: {total_downloaded}/{MAX_DOWNLOADS}")
                
            except Exception as e:
                log_message(f"处理出错: {e}", "ERROR")
            
            # 如果下载失败，增加连续失败计数
            if not download_success:
                consecutive_failures += 1
                if consecutive_failures >= max_consecutive_failures:
                    log_message(f"连续失败 {consecutive_failures} 次，暂停程序", "ERROR")
                    log_message("请检查网络连接或网站状态，按回车键继续...")
                    input()
                    consecutive_failures = 0  # 重置计数
            
            # 短暂等待，避免操作过快
            time.sleep(1)
        
        # 如果重试后仍未成功，记录失败
        if not download_success:
            failed_downloads.append({
                "index": current_index,
                "title": full_title if 'full_title' in locals() else f"结果{idx}",
                "retry_count": retry_count
            })
            log_message(f"文件下载失败，已重试 {retry_count} 次", "ERROR")
    
    # 保存当前的URL记录
    try:
        with open(url_record_file, 'w', encoding='utf-8') as f:
            json.dump(url_records, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log_message(f"保存URL记录失败: {e}", "ERROR")
    
    # 检查是否已达到目标数量
    if total_downloaded >= MAX_DOWNLOADS:
        log_message(f"已达到目标数量 {MAX_DOWNLOADS} 个，任务完成")
        break
    
    # ========== 翻页处理 ==========
    log_message(f"当前页处理完成，尝试翻页...")
    
    try:
        # 查找"下一页"按钮
        next_button = driver.find_element(By.XPATH, "//a[@class='next']")
        
        # 检查按钮是否可用
        if "disabled" in next_button.get_attribute("class"):
            log_message("已到达最后一页，无法继续翻页", "INFO")
            break
        
        # 点击下一页
        log_message("点击下一页按钮...")
        next_button.click()
        page_num += 1
        
        # 等待新页面加载
        log_message("等待新页面加载...")
        time.sleep(5)
        
        # 等待新页面的结果出现
        wait.until(EC.presence_of_element_located((By.XPATH, "//p[@class='tit1']")))
        log_message(f"第 {page_num} 页加载成功")
        
        # 保存检查点
        save_checkpoint()
        
    except NoSuchElementException:
        log_message("未找到'下一页'按钮，可能已到最后一页", "INFO")
        break
    except Exception as e:
        log_message(f"翻页时出错: {e}", "ERROR")
        
        # 尝试其他翻页方法
        try:
            log_message("尝试使用JavaScript翻页...")
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(3)
            
            # 查找是否有分页控件
            pagination_elements = driver.find_elements(By.CLASS_NAME, "pagination")
            if pagination_elements:
                # 尝试找到下一页链接
                next_links = driver.find_elements(By.XPATH, "//a[contains(text(), '下一页') or contains(text(), 'Next')]")
                if next_links:
                    next_links[0].click()
                    time.sleep(5)
                    log_message("通过JavaScript翻页成功")
                else:
                    log_message("未找到下一页链接", "ERROR")
                    break
            else:
                log_message("未找到分页控件", "ERROR")
                break
        except:
            log_message("所有翻页尝试都失败", "ERROR")
            break

# ========== 最终统计和保存 ==========
print(f"\n{'='*70}")
print("任务完成!")
print(f"{'='*70}")
print(f"统计信息:")
print(f"- 目标下载数量: {MAX_DOWNLOADS}")
print(f"- 实际下载数量: {total_downloaded}")
print(f"- 失败数量: {len(failed_downloads)}")
print(f"- 处理页面数: {page_num}")
print(f"- 开始时间: {url_records[0]['timestamp'] if url_records else 'N/A'}")
print(f"- 结束时间: {time.strftime('%Y-%m-%d %H:%M:%S')}")

if failed_downloads:
    print(f"\n失败的文件列表 (共{len(failed_downloads)}个):")
    for i, fail in enumerate(failed_downloads[:20], 1):
        print(f"  {i}. [{fail['index']}] {fail['title'][:50]}... (重试{fail['retry_count']}次)")
    if len(failed_downloads) > 20:
        print(f"  ... 以及另外 {len(failed_downloads)-20} 个失败文件")

# 保存最终的URL记录
print(f"\n正在保存PDF URL记录...")
try:
    with open(url_record_file, 'w', encoding='utf-8') as f:
        json.dump(url_records, f, ensure_ascii=False, indent=2)
    print(f"✓ PDF URL记录已保存到: {url_record_file}")
    
    # 同时保存为文本文件便于查看
    txt_file = url_record_file.replace('.json', '.txt')
    with open(txt_file, 'w', encoding='utf-8') as f:
        f.write("PDF批量下载任务报告\n")
        f.write("="*70 + "\n")
        f.write(f"任务时间: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"搜索关键词: {search_keyword}\n")
        f.write(f"目标数量: {MAX_DOWNLOADS}\n")
        f.write(f"完成数量: {total_downloaded}\n")
        f.write(f"失败数量: {len(failed_downloads)}\n")
        f.write(f"处理页面: {page_num}\n")
        f.write("="*70 + "\n\n")
        
        f.write("下载文件列表:\n")
        for record in url_records:
            if record.get('download_status') == 'success':
                f.write(f"[{record.get('index', 'N/A'):04d}] {record.get('saved_filename', 'N/A')}\n")
                f.write(f"    标题: {record.get('title', '')[:60]}...\n")
                f.write(f"    URL: {record.get('pdf_url', 'N/A')}\n")
                f.write(f"    大小: {record.get('file_size_kb', 0):.1f} KB\n")
                f.write("-"*50 + "\n")
    
    print(f"✓ 文本格式报告已保存到: {txt_file}")
    
except Exception as e:
    print(f"✗ 保存记录失败: {e}")

# 删除检查点文件（任务完成）
if os.path.exists(checkpoint_file):
    os.remove(checkpoint_file)
    print(f"✓ 检查点文件已删除")

print(f"\n所有PDF文件保存到: {save_dir}")
print(f"文件总数: {len([f for f in os.listdir(save_dir) if f.endswith('.pdf')])}")

# 生成文件列表
list_file = os.path.join(save_dir, f"file_list_{timestamp}.txt")
with open(list_file, 'w', encoding='utf-8') as f:
    pdf_files = [file for file in os.listdir(save_dir) if file.endswith('.pdf')]
    pdf_files.sort()
    f.write(f"PDF文件列表 (共{len(pdf_files)}个)\n")
    f.write("="*70 + "\n")
    for pdf_file in pdf_files:
        file_path = os.path.join(save_dir, pdf_file)
        file_size = os.path.getsize(file_path) / 1024  # KB
        f.write(f"{pdf_file} ({file_size:.1f} KB)\n")

print(f"✓ 文件列表已保存到: {list_file}")

# 询问是否打开保存目录
response = input("\n是否要打开保存目录查看文件？(y/n): ")
if response.lower() == 'y':
    try:
        os.startfile(save_dir)
        print(f"已打开目录: {save_dir}")
    except:
        print("无法自动打开目录，请手动访问")

print(f"\n任务完成，按回车键关闭浏览器...")
input()
driver.quit()
