# plugins/scan_disk.py

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

class DiskScanPlugin(ScanPlugin):
    @property
    def name(self):
        return "硬盘"

    def scan(self, wmi_connector):
        data = []
        try:
            for disk in wmi_connector.Win32_DiskDrive():
                data.append({
                    '类别': '硬盘',
                    '品牌': disk.Model.split()[0] if disk.Model else "N/A",
                    '型号': disk.Model or "N/A",
                    '大小': format_bytes(disk.Size),
                    '序列号': disk.SerialNumber.strip() if disk.SerialNumber else "无法获取",
                    '生产日期': 'N/A',
                    '保修查询链接': 'N/A'
                })
        except Exception as e:
            print(f"扫描硬盘失败: {e}")
        return data