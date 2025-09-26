# plugins/scan_peripherals.py

from plugin_interface import ScanPlugin


def filter_devices(devices, blacklist):
    """辅助函数，用于过滤掉通用设备驱动，优先显示具体型号"""
    specific_devices, generic_devices = [], []
    for device in devices:
        description = getattr(device, 'Description', '') or ''
        name = getattr(device, 'Name', '') or ''
        full_description = (description + name).lower()

        is_generic = any(item.lower() in full_description for item in blacklist)

        if is_generic:
            generic_devices.append(device)
        else:
            specific_devices.append(device)

    # 如果有具体设备信息，则返回具体设备；否则返回通用设备信息
    return specific_devices if specific_devices else generic_devices


class PeripheralsScanPlugin(ScanPlugin):
    @property
    def name(self):
        return "键盘和鼠标"

    def scan(self, wmi_connector):
        data = []

        # 键盘黑名单，包含常见通用驱动名称
        keyboard_blacklist = ["HID Keyboard Device", "USB 输入设备", "PS/2"]
        # 鼠标黑名单
        mouse_blacklist = ["HID-compliant mouse", "USB 输入设备"]

        try:
            # 扫描键盘
            keyboards = wmi_connector.Win32_Keyboard()
            for kbd in filter_devices(keyboards, keyboard_blacklist):
                data.append({
                    '类别': '键盘', '品牌': kbd.Name.split()[0], '型号': kbd.Description,
                    '大小': 'N/A', '序列号': 'N/A', '生产日期': 'N/A', '保修查询链接': 'N/A'
                })

            # 扫描鼠标
            pointing_devices = wmi_connector.Win32_PointingDevice()
            for ptr in filter_devices(pointing_devices, mouse_blacklist):
                data.append({
                    '类别': '鼠标', '品牌': ptr.Manufacturer, '型号': ptr.Description,
                    '大小': 'N/A', '序列号': 'N/A', '生产日期': 'N/A', '保修查询链接': 'N/A'
                })
        except Exception as e:
            print(f"扫描外设失败: {e}")

        return data