"""
# -*- coding: utf-8 -*-
# @Time     : 2025/3/24 下午10:11
# @Author   : 代志伟
# @File     : 1.py
# code is far away from bugs with the god animal protecting
    I love animals. They taste delicious.
              ┏┓      ┏┓
            ┏┛┻━━━━━━━━━━━┛┻┓
            ┃     ☃   ┃
            ┃  ┳━━┛  ┗━━┳ ┃
            ┃     ━┻━   ┃
            ┗━┓      ┏━┛
                ┃      ┗━━━┓
                ┃  神兽保佑    ┣┓
                ┃　永无BUG！   ┏┛
                ┗┓┓┏━━━━┳┓┏┛
                  ┃┫┫  ┃┫┫
                  ┗┻┛  ┗┻┛
"""
import json
import re
import random


class ResponseProcess:
    def __init__(self):
        self.id_pattern = re.compile(r'[iI][dD]$')

    def traverse_data(
            self,
            data,
            path=None,
            id_results=None,
            top_level_entries=None,
            nested_entries=None,
            is_top_level=True
    ):
        if path is None:
            path = []
        if id_results is None:
            id_results = []
        if top_level_entries is None:
            top_level_entries = []
        if nested_entries is None:
            nested_entries = []

        try:
            if isinstance(data, dict):
                for k, v in data.items():
                    current_path = path + [f'Key[{k}]']
                    entry = {
                        'path': ' → '.join(current_path),
                        'key': k,
                        'value': v
                    }

                    # 记录顶层非嵌套键值对
                    if is_top_level and not isinstance(v, (dict, list)):
                        top_level_entries.append(entry)

                    # 记录嵌套层级的所有键值对（排除顶层）
                    if not is_top_level and not isinstance(v, (dict, list)):
                        nested_entries.append(entry)

                    # 匹配ID规则的键
                    if self.id_pattern.search(k):
                        id_results.append(entry)

                    # 递归处理子元素
                    self.traverse_data(v, current_path, id_results, top_level_entries, nested_entries, False)

            elif isinstance(data, list):
                for idx, item in enumerate(data):
                    current_path = path + [f'Index[{idx}]']
                    self.traverse_data(item, current_path, id_results, top_level_entries, nested_entries, False)

        except Exception as e:
            print(f"Error at path {' → '.join(path)}: {str(e)}")

        return top_level_entries, nested_entries, id_results

    def load_json(self, json_str):
        data = json.loads(json_str)
        top_level_entries, nested_entries, id_matches = self.traverse_data(data)

        # 提取所有顶层非嵌套键值对
        results = []
        for item in top_level_entries:
            results.append({item['key']: item['value']})

        # === 修改点：优先选取嵌套中以id结尾的键值对 ===
        # 1. 筛选嵌套中以id结尾的键值对
        id_nested = [entry for entry in nested_entries if self.id_pattern.search(entry['key'])]

        # 2. 随机选取最多3个id键值对
        selected_id = []
        try:
            selected_id = random.sample(id_nested, min(3, len(id_nested)))
        except ValueError:
            selected_id = id_nested[:min(3, len(id_nested))]

        # 3. 计算剩余需要补充的数量
        remaining = max(0, 3 - len(selected_id))

        # 4. 从非id的嵌套键中随机选取剩余数量
        non_id_nested = [entry for entry in nested_entries if not self.id_pattern.search(entry['key'])]
        try:
            selected_non_id = random.sample(non_id_nested, remaining)
        except ValueError:
            selected_non_id = non_id_nested[:remaining]

        # 5. 合并结果
        selected_nested = selected_id + selected_non_id

        # 添加嵌套键值对到结果
        for item in selected_nested:
            results.append({item['key']: item['value']})

        return results


if __name__ == '__main__':
    response = ResponseProcess()
    print(response.load_json(
        json_str='{"totle":1, "name": 2,"row":{"updatetime":"2024-08-27 11:06:14","remark":50,"ordNum":5703,"propType":"1","propNum":50000,"trans":"1","reportType":"703","cellName":"[1]现金|[A ]人民币","entName":"1.1 现金"},"msgCode":200,"message":"查询成功"}'
    ))