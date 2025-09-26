# plugin_manager.py (Corrected Version)

import os
import importlib.util
import inspect

# Ensure all plugin interfaces are imported
from plugin_interface import ScanPlugin, ExportPlugin, DiagnosticPlugin, SyncPlugin


class PluginManager:
    """
    Finds, loads, and manages all types of plugins by inspecting their class inheritance.
    """

    def __init__(self, plugin_folder='plugins'):
        self.plugin_folder = plugin_folder
        self.scan_plugins = []
        self.export_plugins = []
        self.diagnostic_plugins = []
        self.sync_plugins = []  # List for sync plugins

    def discover_plugins(self):
        """
        Finds and loads all valid plugins from the plugin directory.
        """
        base_dir = os.path.dirname(os.path.abspath(__file__))
        plugin_path = os.path.join(base_dir, self.plugin_folder)

        if not os.path.isdir(plugin_path):
            print(f"Error: Plugin directory '{plugin_path}' not found.")
            return

        print(f"--- Loading plugins from '{plugin_path}' ---")
        for filename in os.listdir(plugin_path):
            if not filename.endswith('.py') or filename.startswith('__'):
                continue

            module_name = filename[:-3]
            module_path = os.path.join(plugin_path, filename)

            try:
                spec = importlib.util.spec_from_file_location(module_name, module_path)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)

                for name, obj in inspect.getmembers(module):
                    if inspect.isclass(obj) and obj.__module__ == module_name:
                        # Check inheritance against all known interfaces
                        if issubclass(obj, SyncPlugin) and obj is not SyncPlugin:
                            self.sync_plugins.append(obj())
                            print(f"  [Success] Loaded Sync Plugin: {obj.__name__} from {filename}")
                            break
                        elif issubclass(obj, ScanPlugin) and obj is not ScanPlugin:
                            self.scan_plugins.append(obj())
                            print(f"  [Success] Loaded Scan Plugin: {obj.__name__} from {filename}")
                            break
                        elif issubclass(obj, ExportPlugin) and obj is not ExportPlugin:
                            self.export_plugins.append(obj())
                            print(f"  [Success] Loaded Export Plugin: {obj.__name__} from {filename}")
                            break
                        elif issubclass(obj, DiagnosticPlugin) and obj is not DiagnosticPlugin:
                            self.diagnostic_plugins.append(obj())
                            print(f"  [Success] Loaded Diagnostic Plugin: {obj.__name__} from {filename}")
                            break
            except Exception as e:
                print(f"  [Failed] Could not load {filename}: {e}")
        print("--- Plugin loading complete ---")

    def get_scan_plugins(self):
        """Returns a list of all loaded scan plugin instances."""
        return self.scan_plugins

    def get_export_plugins(self):
        """Returns a list of all loaded export plugin instances."""
        return self.export_plugins

    def get_diagnostic_plugins(self):
        """Returns a list of all loaded diagnostic plugin instances."""
        return self.diagnostic_plugins

    # ▼▼▼ THIS IS THE MISSING METHOD ▼▼▼
    def get_sync_plugins(self):
        """Returns a list of all loaded sync plugin instances."""
        return self.sync_plugins