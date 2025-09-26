# plugins/scan_monitor.py

from plugin_interface import ScanPlugin
import wmi


class MonitorScanPlugin(ScanPlugin):
    @property
    def name(self):
        return "显示器信息"

    def scan(self, wmi_connector):
        data = []
        try:
            # 显示器信息在 'wmi' 命名空间下
            c_wmi_monitor = wmi.WMI(namespace="wmi")
            monitors = c_wmi_monitor.WmiMonitorID()

            if not monitors:
                data.append(
                    {'类别': '显示器', '品牌': '无法获取', '型号': '未检测到外部显示器', '大小': 'N/A', '序列号': 'N/A',
                     '生产日期': 'N/A', '保修查询链接': 'N/A'})
                return data

            for mon in monitors:
                # WMI返回的是ASCII码列表，需要转换
                manufacturer = "".join([chr(i) for i in mon.ManufacturerName if i > 0])
                model = "".join([chr(i) for i in mon.UserFriendlyName if i > 0])
                serial = "".join([chr(i) for i in mon.SerialNumberID if i > 0])
                year, week = mon.YearOfManufacture, mon.WeekOfManufacture
                prod_date = f"{year}-W{week}" if year and week else "N/A"

                data.append({
                    '类别': '显示器',
                    '品牌': manufacturer,
                    '型号': model,
                    '大小': 'N/A',  # WMI不直接提供尺寸
                    '序列号': serial,
                    '生产日期': prod_date,
                    '保修查询链接': 'N/A'
                })
        except Exception as e:
            print(f"扫描显示器失败: {e}")
            data.append({'类别': '显示器', '品牌': '获取失败', '型号': str(e), '大小': 'N/A', '序列号': 'N/A',
                         '生产日期': 'N/A', '保修查询链接': 'N/A'})

        return data