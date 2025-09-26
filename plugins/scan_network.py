from plugin_interface import ScanPlugin

class NetworkScanPlugin(ScanPlugin):
    @property
    def name(self):
        return "网络适配器"

    def scan(self, wmi_connector):
        data = []
        try:
            # 筛选出启用了IP且有MAC地址的物理网卡
            for nic in wmi_connector.Win32_NetworkAdapterConfiguration(IPEnabled=True):
                if nic.MACAddress:
                    data.append({
                        '类别': '网卡',
                        '品牌': nic.Description,
                        '型号': f"MAC: {nic.MACAddress}",
                        '大小': 'N/A',
                        '序列号': (nic.IPAddress[0] if nic.IPAddress else "N/A"),
                        '生产日期': 'N/A',
                        '保修查询链接': 'N/A'
                    })
        except Exception as e:
            print(f"扫描网卡失败: {e}")
        return data