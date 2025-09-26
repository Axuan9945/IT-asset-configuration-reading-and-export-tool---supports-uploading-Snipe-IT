# plugins/scan_gpu.py

from plugin_interface import ScanPlugin

class GpuScanPlugin(ScanPlugin):
    @property
    def name(self):
        return "显卡"

    def scan(self, wmi_connector):
        data = []
        try:
            for gpu in wmi_connector.Win32_VideoController():
                data.append({
                    '类别': '显卡',
                    '品牌': gpu.Name.split()[0] if gpu.Name else "N/A",
                    '型号': gpu.Name or "N/A",
                    '大小': 'N/A',
                    '序列号': "N/A",
                    '生产日期': 'N/A',
                    '保修查询链接': 'N/A'
                })
        except Exception as e:
            print(f"扫描显卡失败: {e}")
        return data