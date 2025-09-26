# plugins/scan_os.py

from plugin_interface import ScanPlugin

class OsScanPlugin(ScanPlugin):
    @property
    def name(self):
        return "操作系统"

    def scan(self, wmi_connector):
        data = []
        try:
            os_info = wmi_connector.Win32_OperatingSystem()[0]
            data.append({
                '类别': '操作系统',
                '品牌': 'Microsoft',
                '型号': os_info.Caption,
                '大小': 'N/A',
                '序列号': os_info.SerialNumber, # 系统序列号
                '生产日期': 'N/A',
                '保修查询链接': 'N/A'
            })
        except Exception as e:
            print(f"扫描操作系统失败: {e}")
        return data