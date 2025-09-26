# plugins/scan_cpu.py
from plugin_interface import ScanPlugin

class CPUScanPlugin(ScanPlugin):
    @property
    def name(self):
        return "CPU 信息"

    def scan(self, wmi_connector):
        data = []
        try:
            for cpu in wmi_connector.Win32_Processor():
                data.append({
                    '类别': 'CPU',
                    '品牌': cpu.Manufacturer,
                    '型号': cpu.Name.strip(),
                    '大小': 'N/A',
                    '序列号': cpu.ProcessorId or "无法获取",
                    '生产日期': 'N/A',
                    '保修查询链接': 'N/A'
                })
        except Exception as e:
            print(f"扫描CPU失败: {e}")
            # 即使失败也返回一个条目，让用户知道尝试过
            data.append({'类别': 'CPU', '品牌': '扫描失败', '型号': str(e), '大小': 'N/A', '序列号': 'N/A', '生产日期': 'N/A', '保修查询链接': 'N/A'})
        return data