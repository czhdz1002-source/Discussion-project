import os
import pandas as pd
import warnings
from datetime import datetime
from pathlib import Path

warnings.filterwarnings('ignore')

class PDFProcessor:
    def __init__(self, pdf_folder, output_folder):
        """
        初始化PDF处理器
        
        Args:
            pdf_folder: PDF文件夹路径
            output_folder: TXT输出文件夹路径
        """
        self.pdf_folder = Path(pdf_folder)
        self.output_folder = Path(output_folder)
        self.report_data = []
        
        # 创建输出文件夹
        self.output_folder.mkdir(exist_ok=True, parents=True)
        
        # 创建Excel报告文件夹
        self.report_folder = self.output_folder.parent / "转换报告"
        self.report_folder.mkdir(exist_ok=True)
        
        # Excel报告路径
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.excel_report_path = self.report_folder / f"PDF转换报告_{timestamp}.xlsx"
    
    def detect_pdf_type(self, pdf_path):
        """
        检测PDF类型：文本型、扫描版或混合型
        
        Args:
            pdf_path: PDF文件路径
            
        Returns:
            类型字符串: 'text', 'scanned', 'mixed'
        """
        try:
            import PyPDF2
            
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                text_content = ""
                total_chars = 0
                
                for page_num in range(min(3, len(pdf_reader.pages))):  # 检查前3页
                    page = pdf_reader.pages[page_num]
                    text = page.extract_text()
                    if text:
                        text_content += text
                        total_chars += len(text)
                
                # 如果提取的字符很少，可能是扫描版
                if total_chars < 100:
                    return 'scanned'
                elif total_chars < 500:
                    return 'mixed'  # 可能是混合类型
                else:
                    return 'text'
                    
        except Exception as e:
            print(f"检测PDF类型时出错: {e}")
            return 'unknown'
    
    def extract_text_with_tables(self, pdf_path):
        """
        提取PDF文本和表格
        
        Args:
            pdf_path: PDF文件路径
            
        Returns:
            tuple: (文本内容, 表格数量)
        """
        try:
            import pdfplumber
            
            all_text = ""
            table_count = 0
            
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    # 提取文本
                    text = page.extract_text()
                    if text:
                        all_text += f"=== 第{page_num}页 ===\n"
                        all_text += text + "\n\n"
                    
                    # 提取表格
                    tables = page.extract_tables()
                    if tables:
                        for table_num, table in enumerate(tables, 1):
                            if table:  # 确保表格不为空
                                all_text += f"--- 第{page_num}页 表格{table_num} ---\n"
                                table_count += 1
                                
                                # 处理表格内容
                                for row in table:
                                    # 过滤None值并转换为字符串
                                    row_text = [str(cell) if cell is not None else "" for cell in row]
                                    all_text += "\t".join(row_text) + "\n"
                                all_text += "\n"
            
            return all_text, table_count
            
        except Exception as e:
            print(f"提取PDF内容时出错: {e}")
            return "", 0
    
    def process_pdf(self, pdf_file):
        """
        处理单个PDF文件
        
        Args:
            pdf_file: PDF文件名
            
        Returns:
            dict: 处理结果信息
        """
        pdf_path = self.pdf_folder / pdf_file
        txt_filename = pdf_file.replace('.pdf', '.txt').replace('.PDF', '.txt')
        txt_path = self.output_folder / txt_filename
        
        result_info = {
            '文件名': pdf_file,
            '文件大小(MB)': round(pdf_path.stat().st_size / (1024*1024), 2),
            '转换状态': '',
            'PDF类型': '',
            '提取页数': 0,
            '提取表格数': 0,
            '文本字数': 0,
            '保存路径': str(txt_path),
            '处理时间': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            '备注': ''
        }
        
        try:
            # 1. 检测PDF类型
            pdf_type = self.detect_pdf_type(pdf_path)
            result_info['PDF类型'] = pdf_type
            
            if pdf_type == 'scanned':
                result_info['转换状态'] = '跳过(扫描版)'
                result_info['备注'] = '扫描版PDF，需要OCR处理'
                print(f"跳过扫描版: {pdf_file}")
                return result_info
            
            # 2. 提取文本和表格
            text_content, table_count = self.extract_text_with_tables(pdf_path)
            
            # 3. 统计信息
            import PyPDF2
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                result_info['提取页数'] = len(pdf_reader.pages)
            
            result_info['提取表格数'] = table_count
            result_info['文本字数'] = len(text_content)
            
            if len(text_content.strip()) == 0:
                result_info['转换状态'] = '失败'
                result_info['备注'] = '未提取到文本内容'
                print(f"警告: {pdf_file} 未提取到文本内容")
            else:
                # 4. 保存TXT文件
                with open(txt_path, 'w', encoding='utf-8') as f:
                    f.write(text_content)
                
                result_info['转换状态'] = '成功'
                print(f"成功转换: {pdf_file} (表格: {table_count}个, 字数: {result_info['文本字数']})")
                
        except Exception as e:
            result_info['转换状态'] = '失败'
            result_info['备注'] = f"错误: {str(e)}"
            print(f"转换失败 {pdf_file}: {e}")
        
        return result_info
    
    def process_all_pdfs(self):
        """
        处理所有PDF文件
        """
        print("=" * 60)
        print("PDF批量转换工具")
        print(f"PDF文件夹: {self.pdf_folder}")
        print(f"输出文件夹: {self.output_folder}")
        print("=" * 60)
        
        # 获取所有PDF文件
        pdf_files = [f for f in os.listdir(self.pdf_folder) if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            print("未找到PDF文件！")
            return
        
        print(f"找到 {len(pdf_files)} 个PDF文件\n")
        
        # 处理每个PDF
        for i, pdf_file in enumerate(pdf_files, 1):
            print(f"[{i}/{len(pdf_files)}] 处理: {pdf_file}")
            result = self.process_pdf(pdf_file)
            self.report_data.append(result)
        
        # 生成Excel报告
        self.generate_excel_report()
    
    def generate_excel_report(self):
        """
        生成Excel报告
        """
        if not self.report_data:
            print("没有数据生成报告")
            return
        
        df = pd.DataFrame(self.report_data)
        
        # 重新排列列顺序
        column_order = ['文件名', '文件大小(MB)', 'PDF类型', '转换状态', 
                       '提取页数', '提取表格数', '文本字数', '保存路径', '处理时间', '备注']
        df = df[column_order]
        
        # 保存到Excel
        with pd.ExcelWriter(self.excel_report_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='转换报告', index=False)
            
            # 获取工作簿和工作表
            workbook = writer.book
            worksheet = writer.sheets['转换报告']
            
            # 调整列宽
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"\n{'='*60}")
        print("处理完成！")
        print(f"Excel报告已保存: {self.excel_report_path}")
        
        # 显示统计信息
        self.show_statistics(df)
    
    def show_statistics(self, df):
        """
        显示统计信息
        
        Args:
            df: 数据框
        """
        print("\n统计信息:")
        print(f"总文件数: {len(df)}")
        print(f"成功转换: {len(df[df['转换状态'] == '成功'])}")
        print(f"跳过(扫描版): {len(df[df['转换状态'] == '跳过(扫描版)'])}")
        print(f"转换失败: {len(df[df['转换状态'] == '失败'])}")
        
        if len(df[df['转换状态'] == '成功']) > 0:
            print(f"提取表格总数: {df['提取表格数'].sum()}")
            print(f"提取文本总字数: {df['文本字数'].sum():,}")
    
    def create_summary_report(self):
        """
        创建汇总报告（可选）
        """
        summary_path = self.report_folder / "转换汇总.txt"
        
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write("PDF批量转换汇总报告\n")
            f.write("=" * 50 + "\n")
            f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"PDF文件夹: {self.pdf_folder}\n")
            f.write(f"TXT输出文件夹: {self.output_folder}\n")
            f.write(f"Excel报告: {self.excel_report_path}\n")
            f.write("=" * 50 + "\n\n")
            
            # 统计信息
            df = pd.DataFrame(self.report_data)
            success_count = len(df[df['转换状态'] == '成功'])
            scanned_count = len(df[df['转换状态'] == '跳过(扫描版)'])
            failed_count = len(df[df['转换状态'] == '失败'])
            
            f.write("统计概览:\n")
            f.write(f"  总PDF文件数: {len(df)}\n")
            f.write(f"  成功转换: {success_count}\n")
            f.write(f"  跳过(扫描版): {scanned_count}\n")
            f.write(f"  转换失败: {failed_count}\n")
            
            if success_count > 0:
                f.write(f"  提取表格总数: {df['提取表格数'].sum()}\n")
                f.write(f"  提取文本总字数: {df['文本字数'].sum():,}\n")
            
            # 文件列表
            f.write("\n详细文件列表:\n")
            f.write("-" * 50 + "\n")
            
            for item in self.report_data:
                status_icon = "✓" if item['转换状态'] == '成功' else "○" if item['转换状态'] == '跳过(扫描版)' else "✗"
                f.write(f"{status_icon} {item['文件名']} ({item['文件大小(MB)']}MB) - {item['转换状态']}")
                if item['转换状态'] == '成功':
                    f.write(f" [表格:{item['提取表格数']}, 字数:{item['文本字数']}]")
                f.write("\n")
        
        print(f"汇总报告: {summary_path}")


def get_folder_path():
    """
    获取用户输入的文件夹路径
    """
    while True:
        print("\n请选择输入方式：")
        print("1. 手动输入文件夹路径")
        print("2. 直接粘贴路径（自动去除引号）")
        print("3. 使用默认路径 (D:\\This computer\\desktop\\失信人文件)")
        
        choice = input("请选择 (1/2/3): ").strip()
        
        if choice == '1':
            folder_path = input("请输入PDF文件夹路径: ").strip()
        elif choice == '2':
            folder_path = input("请粘贴文件夹路径: ").strip()
            # 去除可能存在的引号
            folder_path = folder_path.strip('"').strip("'")
        elif choice == '3':
            folder_path = r"D:\This computer\desktop\失信人文件"
            print(f"使用默认路径: {folder_path}")
        else:
            print("无效选择，请重新输入")
            continue
        
        # 验证路径是否存在
        if os.path.exists(folder_path):
            return folder_path
        else:
            print(f"错误：路径不存在 '{folder_path}'，请重新输入")


def main():
    """主函数"""
    print("=" * 60)
    print("欢迎使用PDF批量转换工具")
    print("=" * 60)
    
    # 获取PDF文件夹路径
    pdf_folder = get_folder_path()
    
    # 自动设置输出文件夹
    output_folder = os.path.join(pdf_folder, "TXT输出")
    
    print(f"\nPDF文件夹: {pdf_folder}")
    print(f"TXT输出文件夹: {output_folder}")
    
    # 确认是否继续
    confirm = input("\n确认以上路径无误？(y/n): ").strip().lower()
    if confirm != 'y':
        print("操作已取消")
        return
    
    # 创建处理器并运行
    processor = PDFProcessor(pdf_folder, output_folder)
    processor.process_all_pdfs()
    
    # 创建汇总报告
    processor.create_summary_report()
    
    print("\n" + "=" * 60)
    print("程序执行完毕！")
    input("按回车键退出...")


if __name__ == "__main__":
    main()
