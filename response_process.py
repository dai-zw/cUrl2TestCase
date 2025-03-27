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
import logging

class ResponseProcess:
    def __init__(self):
        self.id_pattern = re.compile(r'[iI][dD]$')
        # 初始化日志记录器
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)

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
            self.logger.debug(f"进入遍历流程，当前路径: {' → '.join(path)}，数据类型: {type(data)}")

            if isinstance(data, dict):
                self.logger.debug(f"开始处理字典，共 {len(data)} 个键")
                for k, v in data.items():
                    current_path = path + [f'Key[{k}]']
                    entry = {
                        'path': ' → '.join(current_path),
                        'key': k,
                        'value': v
                    }

                    # 记录顶层非嵌套键值对
                    if is_top_level and not isinstance(v, (dict, list)):
                        self.logger.debug(f"发现顶层键值对: {k} = {v}")
                        top_level_entries.append(entry)

                    # 记录嵌套层级的所有键值对（排除顶层）
                    if not is_top_level and not isinstance(v, (dict, list)):
                        self.logger.debug(f"发现嵌套键值对: {k} = {v} (路径: {entry['path']})")
                        nested_entries.append(entry)

                    # 匹配ID规则的键
                    if self.id_pattern.search(k):
                        self.logger.info(f"发现ID键: {k}，值: {v} (路径: {entry['path']})")
                        id_results.append(entry)

                    # 递归处理子元素
                    self.traverse_data(v, current_path, id_results, top_level_entries, nested_entries, False)

            elif isinstance(data, list):
                self.logger.debug(f"开始处理列表，共 {len(data)} 个元素")
                for idx, item in enumerate(data):
                    current_path = path + [f'Index[{idx}]']
                    self.traverse_data(item, current_path, id_results, top_level_entries, nested_entries, False)

        except Exception as e:
            self.logger.error(f"遍历过程中发生异常，路径: {' → '.join(path)}，错误信息: {str(e)}", exc_info=True)

        return top_level_entries, nested_entries, id_results

    def load_json(self, json_str):
        self.logger.info("开始解析JSON字符串")
        try:
            data = json.loads(json_str)
            self.logger.debug("JSON解析成功，开始遍历数据结构")
        except json.JSONDecodeError as e:
            self.logger.error(f"JSON解析失败: {str(e)}")
            return []

        top_level_entries, nested_entries, id_matches = self.traverse_data(data)
        self.logger.info(f"数据遍历完成 - 顶层条目: {len(top_level_entries)}，嵌套条目: {len(nested_entries)}，ID匹配项: {len(id_matches)}")

        # 提取所有顶层非嵌套键值对
        results = []
        self.logger.debug("开始处理顶层条目")
        for item in top_level_entries:
            results.append({item['key']: item['value']})
        self.logger.debug(f"已添加{len(top_level_entries)}个顶层条目到结果集")

        # === 修改点：优先选取嵌套中以id结尾的键值对 ===
        id_nested = [entry for entry in nested_entries if self.id_pattern.search(entry['key'])]
        non_id_nested = [entry for entry in nested_entries if not self.id_pattern.search(entry['key'])]
        self.logger.debug(f"筛选嵌套条目 - ID相关: {len(id_nested)}，非ID: {len(non_id_nested)}")

        # 随机选取最多3个id键值对
        selected_id = []
        try:
            sample_size = min(3, len(id_nested))
            selected_id = random.sample(id_nested, sample_size)
            self.logger.info(f"从{len(id_nested)}个ID嵌套项中随机选取{sample_size}个")
        except ValueError as e:
            self.logger.warning(f"ID嵌套项不足，实际选取{len(id_nested)}个: {str(e)}")
            selected_id = id_nested[:min(3, len(id_nested))]

        # 计算剩余需要补充的数量
        remaining = max(0, 3 - len(selected_id))
        self.logger.debug(f"需要补充的非ID条目数量: {remaining}")

        # 从非ID中随机选取剩余数量
        selected_non_id = []
        try:
            if remaining > 0:
                selected_non_id = random.sample(non_id_nested, remaining)
                self.logger.info(f"从{len(non_id_nested)}个非ID项中补充选取{remaining}个")
        except ValueError as e:
            self.logger.warning(f"非ID嵌套项不足，实际选取{len(non_id_nested)}个: {str(e)}")
            selected_non_id = non_id_nested[:remaining]

        # 合并结果
        selected_nested = selected_id + selected_non_id
        self.logger.debug(f"最终选取嵌套条目 - ID: {len(selected_id)}，非ID: {len(selected_non_id)}")

        # 添加嵌套键值对到结果
        self.logger.debug("开始添加嵌套条目到结果集")
        for item in selected_nested:
            results.append({item['key']: item['value']})
        self.logger.info(f"处理完成，最终结果集包含{len(results)}个条目")

        return results


if __name__ == '__main__':
    response = ResponseProcess()
    print(response.load_json(
        json_str='{"totle":1, "name": 2,"row":{"updatetime":"2024-08-27 11:06:14","remark":50,"ordNum":5703,"propType":"1","propNum":50000,"trans":"1","reportType":"703","cellName":"[1]现金|[A ]人民币","entName":"1.1 现金"},"msgCode":200,"message":"查询成功"}'
    ))