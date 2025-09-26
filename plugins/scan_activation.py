# plugins/scan_activation.py

from plugin_interface import ScanPlugin


class ActivationScanPlugin(ScanPlugin):
    @property
    def name(self):
        return "系统激活状态"

    def scan(self, wmi_connector):
        data = []
        status = "未激活或无法确定"
        try:
            for product in wmi_connector.SoftwareLicensingProduct():
                # 寻找包含 "Windows" 描述且有部分产品密钥的授权信息
                if product.Description and "windows" in product.Description.lower() and product.PartialProductKey:
                    # LicenseStatus=1 表示已授权
                    if product.LicenseStatus == 1:
                        status = "已激活"
                        break
        except Exception as e:
            print(f"扫描激活状态失败: {e}")
            status = f"查询失败: {e}"

        data.append({
            '类别': '系统激活状态',
            '品牌': 'N/A',
            '型号': status,
            '大小': 'N/A',
            '序列号': 'N/A',
            '生产日期': 'N/A',
            '保修查询链接': 'N/A'
        })
        return data