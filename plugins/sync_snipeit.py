# plugins/sync_snipeit.py

import requests
import json
from plugin_interface import SyncPlugin


class SnipeITSyncPlugin(SyncPlugin):
    def __init__(self):
        self.name = "同步到 Snipe-IT"
        self.icon_name = "sync"
        self.headers = {}
        self.base_url = ""

    def _determine_active_url(self, worker, internal_url, external_url):
        if internal_url:
            worker.log_message.emit(f"  -> 正在尝试连接内网URL: {internal_url}...")
            try:
                response = requests.get(f"{internal_url.rstrip('/')}/api/v1/statuslabels", headers=self.headers,
                                        timeout=2)
                response.raise_for_status()
                worker.log_message.emit("  -> ✅ 内网URL连接成功，将使用此地址。")
                return internal_url
            except requests.exceptions.RequestException as e:
                worker.log_message.emit(f"  -> ⚠️ 内网URL连接失败: {e}")
        if external_url:
            worker.log_message.emit(f"  -> 正在尝试连接外网URL: {external_url}...")
            try:
                response = requests.get(f"{external_url.rstrip('/')}/api/v1/statuslabels", headers=self.headers,
                                        timeout=5)
                response.raise_for_status()
                worker.log_message.emit("  -> ✅ 外网URL连接成功，将使用此地址。")
                return external_url
            except requests.exceptions.RequestException as e:
                worker.log_message.emit(f"  -> ❌ 外网URL连接失败: {e}")
        return None

    def _api_request(self, worker, method, endpoint, payload=None):
        url = f"{self.base_url.rstrip('/')}/api/v1/{endpoint.lstrip('/')}"
        try:
            if method.upper() == 'GET':
                response = requests.get(url, headers=self.headers, params=payload, timeout=10)
            elif method.upper() == 'POST':
                response = requests.post(url, headers=self.headers, data=json.dumps(payload), timeout=10)
            else:
                raise NotImplementedError(f"不支持的请求方法: {method}")
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            worker.log_message.emit(f"  -> ❌ API 请求失败: {e}")
            return None

    def _get_or_create(self, worker, search_name, endpoint, creation_payload=None):
        worker.log_message.emit(f"  -> 正在查询 {search_name}...")
        response_data = self._api_request(worker, 'GET', endpoint, payload={'search': search_name})
        if response_data and response_data.get('total', 0) > 0:
            for item in response_data['rows']:
                if item.get('name', '').lower() == search_name.lower():
                    worker.log_message.emit(f"  -> ✅ 已找到 '{search_name}' (ID: {item['id']})")
                    return item['id']
        worker.log_message.emit(f"  -> '{search_name}' 不存在，正在创建...")
        payload_to_create = creation_payload if creation_payload is not None else {'name': search_name}
        creation_data = self._api_request(worker, 'POST', endpoint, payload=payload_to_create)
        if creation_data and creation_data.get('status') == 'success':
            new_id = creation_data.get('payload', {}).get('id')
            worker.log_message.emit(f"  -> ✅ 成功创建 '{search_name}' (新 ID: {new_id})")
            return new_id
        worker.log_message.emit(f"  -> ❌ 创建 '{search_name}' 失败。")
        return None

    def sync(self, worker, data: list, config: dict):
        log_callback = worker.log_message.emit
        api_key = config.get('key')
        internal_url = config.get('internal_url')
        external_url = config.get('external_url')

        if not api_key:
            log_callback("❌ 错误：请先在主页配置中填写 Snipe-IT API 密钥。")
            return
        if not internal_url and not external_url:
            log_callback("❌ 错误：请至少填写一个内网或外网URL。")
            return

        self.headers = {"Authorization": f"Bearer {api_key}", "Accept": "application/json",
                        "Content-Type": "application/json"}
        self.base_url = self._determine_active_url(worker, internal_url, external_url)
        if not self.base_url:
            log_callback("❌ 错误：内网和外网URL都无法连接，同步任务中止。")
            return

        log_callback(f"--- 开始同步资产到 Snipe-IT ({self.base_url}) ---")
        CATEGORY_ID_MAP = {'台式机': 1, '笔记本': 2, '显示器': 3}
        main_assets = [item for item in data if
                       item.get('类别') == '主板/整机' and item.get('序列号') and item.get('序列号') != 'N/A']

        for asset_data in main_assets:
            serial = asset_data.get('序列号')
            manufacturer_name = asset_data.get('品牌', '未知制造商')
            model_name = asset_data.get('型号', '未知型号')
            category_name_in_snipeit = '台式机'
            log_callback(f"\n--- 正在处理序列号: {serial} ---")

            manufacturer_id = self._get_or_create(worker, manufacturer_name, 'manufacturers')
            if not manufacturer_id: continue

            category_id = CATEGORY_ID_MAP.get(category_name_in_snipeit)
            if not category_id:
                log_callback(f"  -> ❌ 错误：未在 CATEGORY_ID_MAP 中配置 '{category_name_in_snipeit}' 的ID。")
                continue

            model_payload = {'name': model_name, 'category_id': category_id, 'manufacturer_id': manufacturer_id}
            model_id = self._get_or_create(worker, model_name, 'models', creation_payload=model_payload)
            if not model_id: continue

            asset_payload = {
                "model_id": model_id, "serial": serial,
                "name": asset_data.get('资产名称', f"{manufacturer_name} {model_name}"),
                "status_id": 2, "asset_tag": asset_data.get('资产标签', serial)
            }

            existing_asset = self._api_request(worker, 'GET', f"hardware/byserial/{serial}")
            if existing_asset and existing_asset.get('total', 0) > 0:
                asset_id = existing_asset['rows'][0]['id']
                log_callback(f"  -> ✅ 资产已存在 (ID: {asset_id})，未来可在此处执行更新操作。")
            else:
                log_callback(f"  -> 资产不存在，正在创建...")
                creation_result = self._api_request(worker, 'POST', 'hardware', payload=asset_payload)
                if creation_result and creation_result.get('status') == 'success':
                    log_callback(f"  -> ✅ 成功在 Snipe-IT 中创建新资产！")
                else:
                    log_callback(f"  -> ❌ 创建资产失败。")
        log_callback("\n--- 所有资产同步任务完成 ---")