# plugins/export_pdf.py (最终修正版 - 适配您的数据结构)

import os
from plugin_interface import ExportPlugin
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# --- 字体查找辅助函数 ---
def find_system_font(font_names: list) -> str | None:
    font_dir = os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts")
    if not os.path.isdir(font_dir):
        return None
    for font_name in font_names:
        font_path = os.path.join(font_dir, font_name)
        if os.path.exists(font_path):
            return font_path
    return None


# --- PDF 导出插件类 ---
class PDFExportPlugin(ExportPlugin):
    def __init__(self):
        self.name = "导出为 PDF"
        self.file_extension = "pdf"
        self.file_filter = "PDF 文件 (*.pdf)"
        self.icon_name = "pdf"
        self.preferred_fonts = ['msyh.ttf', 'simhei.ttf', 'deng.ttf', 'simsun.ttc']
        self.pdf_font_name = 'Chinese-Font'

    def export(self, data, output_path, header_text, log_callback):
        try:
            log_callback("  -> 正在自动查找系统中可用的中文字体...")
            font_path = find_system_font(self.preferred_fonts)

            if not font_path:
                error_msg = "导出失败：在系统中找不到任何可用的中文字体。"
                log_callback(f"  -> 错误：{error_msg}")
                return error_msg

            log_callback(f"  -> 成功找到字体: {os.path.basename(font_path)}")
            pdfmetrics.registerFont(TTFont(self.pdf_font_name, font_path))

            c = canvas.Canvas(output_path, pagesize=letter)
            width, height = letter
            y_position = height - 40

            if header_text:
                c.setFont(self.pdf_font_name, 16)
                c.drawCentredString(width / 2.0, y_position, header_text)
                y_position -= 40

            last_category = None
            for item in data:
                # ▼▼▼ 核心修改逻辑开始 ▼▼▼

                # 检查是否需要换页
                if y_position < 60:
                    c.showPage()
                    y_position = height - 40
                    last_category = None  # 换页后重新打印类别标题

                current_category = item.get('类别', '未知类别')

                # 如果是新的类别，打印一个大的类别标题
                if current_category != last_category:
                    y_position -= 20  # 与上一个类别拉开间距
                    c.setFont(self.pdf_font_name, 14)
                    c.drawString(50, y_position, f"--- {current_category} ---")
                    y_position -= 25
                    last_category = current_category

                # 遍历这条数据的所有键值对
                for key, value in item.items():
                    # 我们已经打印过大的类别标题了，所以跳过'类别'这个键
                    if key == '类别':
                        continue

                    # 检查是否需要换页
                    if y_position < 40:
                        c.showPage()
                        y_position = height - 40
                        # 换页后把大标题也重新打上
                        c.setFont(self.pdf_font_name, 14)
                        c.drawString(50, y_position, f"--- {current_category} ---")
                        y_position -= 25

                    # 将键作为“项目”，值作为“详细信息”打印出来
                    c.setFont(self.pdf_font_name, 10)
                    line = f"{key}: {value}"
                    c.drawString(70, y_position, line)
                    y_position -= 20  # 减小行距

                # ▲▲▲ 核心修改逻辑结束 ▲▲▲

            c.save()
            log_callback(f"✅ PDF 文件已成功保存到: {output_path}")
            return output_path

        except Exception as e:
            error_msg = f"导出PDF时发生严重错误: {e}"
            log_callback(f"❌ {error_msg}")
            return error_msg


# 不要忘记这一行
plugin_class = PDFExportPlugin