import csv
from plugin_interface import ExportPlugin


class CSVExportPlugin(ExportPlugin):
    @property
    def name(self):
        return "导出为 CSV"

    @property
    def file_extension(self):
        return ".csv"

    @property
    def file_filter(self):
        return "CSV (逗号分隔) (*.csv)"

    @property
    def icon_name(self):
        return "csv"  # 确保 icons 文件夹里有一个 csv.png 图标

    def export(self, data, file_path, header_text, log_callback):
        log_callback("  -> 开始生成 CSV 数据...")
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                fieldnames = ['类别', '品牌', '型号', '大小', '序列号', '生产日期', '保修查询链接']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

                writer.writeheader()
                for row in data:
                    writer.writerow(row)
            log_callback("  -> CSV 文件写入完成。")
            return file_path
        except Exception as e:
            log_callback(f"  -> 导出 CSV 失败: {e}")
            return None