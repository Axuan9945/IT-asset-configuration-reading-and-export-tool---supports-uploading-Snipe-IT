# plugins/diagnostic_health_check.py (功能完整版)

import ctypes
import datetime
import socket
import wmi
from plugin_interface import DiagnosticPlugin

# 尝试导入可选的库，如果失败则优雅地处理
try:
    import psutil
except ImportError:
    psutil = None

try:
    import ping3

    ping3.EXCEPTIONS = True
except ImportError:
    ping3 = None


class HealthCheckDiagnosticPlugin(DiagnosticPlugin):
    """
    一个功能全面的系统健康检查插件，涵盖系统基础、性能、Windows健康、硬件状态和网络。
    """

    def __init__(self):
        self.wmi_conn = None
        self.wmi_temp_conn = None

    @property
    def name(self):
        return "系统综合诊断"

    def run_diagnostic(self) -> list:
        """
        执行所有诊断任务并返回结果列表。
        """
        results = []

        # 建立WMI连接，供后续检查使用
        try:
            self.wmi_conn = wmi.WMI()
            # 尝试连接到用于获取温度的特殊WMI命名空间
            self.wmi_temp_conn = wmi.WMI(namespace="root\\wmi")
        except Exception as e:
            results.append(
                {'task': 'WMI服务连接', 'status': '失败', 'message': f'无法连接到WMI服务，部分诊断无法执行: {e}'})
            # 如果核心的WMI连接失败，则没有必要继续
            return results

        # --- 按类别执行所有诊断任务 ---
        self._check_system_basics(results)
        self._check_performance(results)
        self._check_windows_health(results)
        self._check_hardware_status(results)
        self._check_network(results)

        return results

    # --- 1. 系统基础诊断 ---
    def _check_system_basics(self, results: list):
        # 检查管理员权限
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
            if is_admin:
                results.append({'task': '权限检查', 'status': '正常', 'message': '程序正以管理员权限运行。'})
            else:
                results.append({'task': '权限检查', 'status': '警告',
                                'message': '程序未以管理员权限运行，可能无法获取所有硬件信息。'})
        except Exception as e:
            results.append({'task': '权限检查', 'status': '错误', 'message': str(e)})

        # 检查系统运行时长
        try:
            os_info = self.wmi_conn.Win32_OperatingSystem()[0]
            last_boot_str = os_info.LastBootUpTime.split('.')[0]
            last_boot_time = datetime.datetime.strptime(last_boot_str, "%Y%m%d%H%M%S")
            uptime = datetime.datetime.now() - last_boot_time
            days, hours, minutes = uptime.days, uptime.seconds // 3600, (uptime.seconds // 60) % 60
            uptime_str = f"{days}天 {hours}小时 {minutes}分钟"
            results.append({'task': '系统运行时长', 'status': '信息', 'message': f'系统已连续运行: {uptime_str}'})
        except Exception as e:
            results.append({'task': '系统运行时长', 'status': '错误', 'message': str(e)})

    # --- 2. 性能诊断 ---
    def _check_performance(self, results: list):
        if not psutil:
            results.append({'task': '性能诊断', 'status': '跳过', 'message': 'psutil 库未安装，无法执行此项检查。'})
            return

        # CPU和内存占用
        cpu_usage = psutil.cpu_percent(interval=1)
        cpu_status = '警告' if cpu_usage > 90 else '正常'
        results.append({'task': 'CPU 总体使用率', 'status': cpu_status, 'message': f'{cpu_usage}%'})

        mem = psutil.virtual_memory()
        mem_status = '警告' if mem.percent > 90 else '正常'
        results.append({'task': '内存使用率', 'status': mem_status,
                        'message': f'{mem.percent}% (已用 {mem.used / 1024 ** 3:.2f} GB / 共 {mem.total / 1024 ** 3:.2f} GB)'})

        # 高资源消耗进程
        try:
            processes = [p for p in psutil.process_iter(['name', 'cpu_percent']) if p.info['cpu_percent'] is not None]
            top_processes = sorted(processes, key=lambda p: p.info['cpu_percent'], reverse=True)
            top_cpu_str = ", ".join([f"{p.info['name']}({p.info['cpu_percent']:.1f}%)" for p in top_processes[:3]])
            results.append({'task': 'CPU占用最高的进程', 'status': '信息', 'message': top_cpu_str})
        except Exception as e:
            results.append({'task': 'CPU占用最高的进程', 'status': '错误', 'message': f'无法获取进程列表: {e}'})

    # --- 3. Windows 系统健康 ---
    def _check_windows_health(self, results: list):
        # 检查关键服务
        critical_services = {'Spooler': '打印服务', 'wuauserv': '更新服务', 'BFE': '防火墙服务'}
        for service_name, display_name in critical_services.items():
            try:
                service = self.wmi_conn.Win32_Service(Name=service_name)[0]
                status = '正常' if service.State == 'Running' else '警告'
                message = f'状态: {service.State}'
                results.append({'task': f'服务 ({display_name})', 'status': status, 'message': message})
            except IndexError:
                results.append({'task': f'服务 ({display_name})', 'status': '失败', 'message': '未找到该服务。'})

        # 检查系统错误日志
        try:
            yesterday = datetime.datetime.now() - datetime.timedelta(days=1)
            query_date = yesterday.strftime("%Y%m%d%H%M%S")
            errors = self.wmi_conn.Win32_NTLogEvent(Logfile='System', EventType=1, TimeGenerated=f">{query_date}")
            status = '警告' if len(errors) > 10 else '正常'
            results.append(
                {'task': '系统错误日志(24h)', 'status': status, 'message': f'发现 {len(errors)} 个严重错误。'})
        except Exception as e:
            results.append({'task': '系统错误日志(24h)', 'status': '错误', 'message': str(e)})

    # --- 4. 硬件状态诊断 ---
    def _check_hardware_status(self, results: list):
        # 硬盘健康
        try:
            for drive in self.wmi_conn.Win32_DiskDrive():
                status = '正常' if drive.Status == "OK" else '警告'
                results.append({'task': f"硬盘健康 ({drive.Caption})", 'status': status,
                                'message': f'S.M.A.R.T. 状态: {drive.Status}'})
        except Exception as e:
            results.append({'task': '硬盘健康 (S.M.A.R.T.)', 'status': '错误', 'message': str(e)})

        # 电池健康
        try:
            batteries = self.wmi_conn.Win32_Battery()
            if batteries:
                health = (batteries[0].FullChargeCapacity / batteries[0].DesignCapacity) * 100
                status = '警告' if health < 80 else '正常'
                results.append({'task': '电池健康度', 'status': status, 'message': f'当前约为 {health:.0f}%'})
        except Exception:
            pass  # 没有电池或查询失败，静默跳过

        # CPU温度
        try:
            if self.wmi_temp_conn:
                temp_info = self.wmi_temp_conn.MSAcpi_ThermalZoneTemperature()[0]
                temp_c = (temp_info.CurrentTemperature / 10.0) - 273.15
                status = '警告' if temp_c > 90 else '正常'
                results.append({'task': 'CPU 温度', 'status': status, 'message': f'{temp_c:.1f} °C'})
        except Exception:
            results.append({'task': 'CPU 温度', 'status': '信息', 'message': '无法从此设备获取温度读数。'})

    # --- 5. 网络连接诊断 ---
    def _check_network(self, results: list):
        if not ping3:
            results.append({'task': '网络诊断', 'status': '跳过', 'message': 'ping3 库未安装。'})
            return

        gateway = self._get_default_gateway()
        if gateway:
            self._perform_ping(f"内网网关 ({gateway})", gateway, results)
        else:
            results.append({'task': '内网网关', 'status': '警告', 'message': '未能自动找到内网网关地址。'})

        target_host = "www.baidu.com"
        try:
            ip = socket.gethostbyname(target_host)
            results.append({'task': 'DNS解析', 'status': '正常', 'message': f'{target_host} 成功解析到 IP: {ip}'})
        except Exception as e:
            results.append({'task': 'DNS解析', 'status': '失败', 'message': f'无法解析 {target_host}，可能无法上网。'})

        self._perform_ping(f"外网连接 ({target_host})", target_host, results)

    def _get_default_gateway(self) -> str | None:
        try:
            routes = self.wmi_conn.Win32_IP4RouteTable(Destination='0.0.0.0', Mask='0.0.0.0')
            return routes[0].NextHop if routes else None
        except Exception:
            return None

    def _perform_ping(self, task_name: str, target: str, results: list):
        try:
            delay = ping3.ping(target, unit='ms', timeout=2)
            if delay is not None:
                results.append({'task': task_name, 'status': '正常', 'message': f'连接成功，延迟: {delay:.2f} ms'})
            else:
                results.append({'task': task_name, 'status': '失败', 'message': '连接超时，目标无响应。'})
        except PermissionError:
            results.append({'task': task_name, 'status': '错误', 'message': '权限不足，执行Ping需要管理员权限。'})
        except Exception as e:
            results.append({'task': task_name, 'status': '失败', 'message': f'连接失败，错误: {e}'})

# 智能插件管理器可以自动发现这个类