# plugins/scan_motherboard.py

from plugin_interface import ScanPlugin

class MotherboardScanPlugin(ScanPlugin):
    @property
    def name(self):
        return "主板/整机"

    def scan(self, wmi_connector):
        data = []
        try:
            for board in wmi_connector.Win32_BaseBoard():
                manufacturer = board.Manufacturer
                serial_number = board.SerialNumber.strip() if board.SerialNumber else ""
                warranty_link = 'N/A'

                # 只有在获取到有效序列号时才生成链接
                if serial_number and "serial" not in serial_number.lower() and "none" not in serial_number.lower():
                    mfg_lower = manufacturer.lower()
                    if 'dell' in mfg_lower:
                        warranty_link = f"https://www.dell.com/support/home/en-sg/product-support/servicetag/{serial_number}/overview"
                    elif 'hewlett-packard' in mfg_lower or 'hp' in mfg_lower:
                        warranty_link = f"https://support.hp.com/sg-en/checkwarranty/search?q={serial_number}"
                    elif 'lenovo' in mfg_lower:
                        warranty_link = f"https://pcsupport.lenovo.com/sg-en/search?query={serial_number}"

                data.append({
                    '类别': '主板/整机',
                    '品牌': manufacturer,
                    '型号': board.Product,
                    '大小': 'N/A',
                    '序列号': serial_number,
                    '生产日期': 'N/A',
                    '保修查询链接': warranty_link
                })
        except Exception as e:
            print(f"扫描主板失败: {e}")
        return data