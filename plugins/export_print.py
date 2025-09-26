# plugins/export_print.py

import os
import tempfile
import uuid
from plugin_interface import ExportPlugin
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# --- 字体查找辅助函数 ---
def find_system_font(font_names: list) -> str | None:
    """
    在 Windows 字体目录中按顺序查找可用的字体文件。
    """
    font_dir = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts")
    if not os.path.isdir(font_dir):
        return None

    for font_name in font_names:
        font_path = os.path.join(font_dir, font_name)
        if os.path.exists(font_path):
            return font_path
    return None


# --- 打印导出插件类 ---
class PrintExportPlugin(ExportPlugin):
    def __init__(self):
        self.name = "打印报告"
        self.file_extension = ""
        self.file_filter = ""
        self.icon_name = "print"
        self.preferred_fonts = ['msyh.ttf', 'simhei.ttf', 'deng.ttf', 'simsun.ttc']
        self.pdf_font_name = 'Chinese-Font-For-Print'

    def export(self, data, output_path, header_text, log_callback, printer_name):
        # 注意：此方法会接收所有参数，但 output_path 不会被使用
        try:
            log_callback(f"  -> 正在为打印机 '{printer_name}' 准备报告...")

            # 1. 自动查找字体
            font_path = find_system_font(self.preferred_fonts)
            if not font_path:
                error_msg = "打印失败：在系统中找不到任何可用的中文字体。"
                log_callback(f"  -> 错误：{error_msg}")
                return error_msg

            log_callback(f"  -> 成功找到字体: {os.path.basename(font_path)}")
            pdfmetrics.registerFont(TTFont(self.pdf_font_name, font_path))

            # 2. 创建临时PDF文件
            temp_dir = tempfile.gettempdir()
            temp_filename = f"asset_report_print_{uuid.uuid4().hex[:8]}.pdf"
            temp_pdf_path = os.path.join(temp_dir, temp_filename)
            log_callback(f"  -> 正在生成临时PDF文件用于打印: {temp_filename}")

            # 3. 生成PDF内容 (此逻辑与 export_pdf.py 完全相同)
            c = canvas.Canvas(temp_pdf_path, pagesize=letter)
            width, height = letter
            y_position = height - 40

            if header_text:
                c.setFont(self.pdf_font_name, 16)
                c.drawCentredString(width / 2.0, y_position, header_text)
                y_position -= 40

            last_category = None
            for item in data:
                if y_position < 60:
                    c.showPage()
                    y_position = height - 40
                    last_category = None

                current_category = item.get('类别', '未知类别')

                if current_category != last_category:
                    y_position -= 20
                    c.setFont(self.pdf_font_name, 14)
                    c.drawString(50, y_position, f"--- {current_category} ---")
                    y_position -= 25
                    last_category = current_category

                for key, value in item.items():
                    if key == '类别':
                        continue

                    if y_position < 40:
                        c.showPage()
                        y_position = height - 40
                        c.setFont(self.pdf_font_name, 14)
                        c.drawString(50, y_position, f"--- {current_category} ---")
                        y_position -= 25

                    c.setFont(self.pdf_font_name, 10)
                    line = f"{key}: {value}"
                    c.drawString(70, y_position, line)
                    y_position -= 20

            c.save()

            # 4. 自动打开生成的临时PDF
            try:
                os.startfile(temp_pdf_path)
                log_callback("  -> 临时PDF文件已创建并打开，请手动打印。")
            except OSError as e:
                error_msg = f"无法自动打开PDF文件。请手动前往临时文件夹查找并打印。\n路径: {temp_pdf_path}\n错误: {e}"
                log_callback(f"  -> 错误：{error_msg}")
                return error_msg

            # 5. 返回特殊字典，告知主程序打印操作已启动
            return {'action': 'print', 'path': temp_pdf_path}

        except Exception as e:
            error_msg = f"准备打印时发生严重错误: {e}"
            log_callback(f"❌ {error_msg}")
            return error_msg

# 智能插件管理器可以自动发现这个类，所以不需要 'plugin_class = ...'