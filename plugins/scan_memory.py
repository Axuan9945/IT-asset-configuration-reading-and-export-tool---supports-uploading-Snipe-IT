# plugins/scan_memory.py

from plugin_interface import ScanPlugin

def format_bytes(byte_size):
    """辅助函数，将字节转换为GB/MB等"""
    try:
        byte_size = int(byte_size)
    except (ValueError, TypeError):
        return "N/A"
    power = 1024
    n = 0
    power_labels = {0: 'B', 1: 'KB', 2: 'MB', 3: 'GB', 4: 'TB'}
    while byte_size >= power and n < len(power_labels) - 1:
        byte_size /= power
        n += 1
    return f"{byte_size:.2f} {power_labels[n]}"

class MemoryScanPlugin(ScanPlugin):
    @property
    def name(self):
        return "内存条"

    def scan(self, wmi_connector):
        data = []
        try:
            for memory in wmi_connector.Win32_PhysicalMemory():
                data.append({
                    '类别': '内存',
                    '品牌': memory.Manufacturer or "N/A",
                    '型号': memory.PartNumber.strip() if memory.PartNumber else "N/A",
                    '大小': format_bytes(memory.Capacity),
                    '序列号': memory.SerialNumber or "N/A",
                    '生产日期': 'N/A',
                    '保修查询链接': 'N/A'
                })
        except Exception as e:
            print(f"扫描内存失败: {e}")
        return data