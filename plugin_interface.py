# plugin_interface.py

from abc import ABC, abstractmethod

class ScanPlugin(ABC):
    @abstractmethod
    def scan(self, wmi_instance) -> list:
        pass

class ExportPlugin(ABC):
    @abstractmethod
    def export(self, data: list, output_path: str, header_text: str, log_callback, printer_name: str):
        pass

class DiagnosticPlugin(ABC):
    @abstractmethod
    def run_diagnostic(self) -> list:
        pass

class SyncPlugin(ABC):
    @abstractmethod
    def sync(self, worker, data: list, config: dict):
        pass