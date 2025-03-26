"""2025/3/22版本：
1.cUrl命令.xlsx中新增请求头字段，以请求方式开头，然后是接口路径，然后是请求头信息 √
2.需要从请求头字段里，获取请求方式以及content-type/Accept √
3.从正式的response里，获取所有单层的键值对，另外从多层键值对里匹配*id的key，如果没有则随机获取3个键值对作为断言 √
"""

import re
import json
from urllib.parse import urlparse
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill
from collections import OrderedDict
from response_process import ResponseProcess
from urllib.parse import unquote

# # 登录接口固定header
# LOGIN_HEADER_FIXED = ''

response_process = ResponseProcess()

# ---------------------- 新增解码函数 ----------------------
def decode_url_encoded(value):
    """安全解码URL编码值"""
    try:
        return unquote(value)
    except Exception as e:
        print(f"解码失败：{value}，错误：{str(e)}")
        return value

def parse_query_with_decode(query_str):
    """解析并解码查询参数"""
    decoded_query = []
    if query_str:
        pairs = query_str.split('&')
        for pair in pairs:
            if '=' in pair:
                key, val = pair.split('=', 1)
                decoded_key = decode_url_encoded(key)
                decoded_val = decode_url_encoded(val)
                decoded_query.append(f"{decoded_key}={decoded_val}")
            else:
                decoded_query.append(decode_url_encoded(pair))
    return '@@'.join(decoded_query)

def decode_nested_values(data):
    """递归解码嵌套结构中的URL编码值"""
    if isinstance(data, dict):
        return {decode_url_encoded(k): decode_nested_values(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [decode_nested_values(item) for item in data]
    elif isinstance(data, str):
        return decode_url_encoded(data)
    else:
        return data

# ---------------------- 核心解析逻辑 ----------------------
def parse_curl(curl_command):
    """从cURL命令解析协议、路径、查询参数和请求体"""
    url_match = re.search(r"'(https?://[^']+)'", curl_command) or re.search(r'"(https?://[^"]+)"', curl_command)
    url = url_match.group(1) if url_match else ''
    parsed_url = urlparse(url)

    data_match = re.search(r"--data-raw '(.*?)'", curl_command) or re.search(r"-d '(.*?)'", curl_command)
    data = data_match.group(1) if data_match else ''

    # 新增：解码查询参数
    if parsed_url:
        decoded_query = parse_query_with_decode(parsed_url.query)
    else:
        decoded_query = None

    # 新增：解码请求体数据
    decoded_data = data
    try:
        if 'application/json' in curl_command.lower() and data:
            # 深度解码嵌套的URL编码值
            json_data = json.loads(data)
            decoded_data = json.dumps(decode_nested_values(json_data), ensure_ascii=False)

        elif 'application/x-www-form-urlencoded' in curl_command.lower() and data:
            # 专用表单数据处理
            decoded_pairs = []
            pairs = data.split('&')
            for pair in pairs:
                if '=' in pair:
                    key, val = pair.split('=', 1)
                    decoded_pairs.append(f"{unquote(key)}={unquote(val)}")
                else:
                    decoded_pairs.append(unquote(pair))
            decoded_data = '&'.join(decoded_pairs)

    except Exception as e:
        print(f"请求体解码失败：{str(e)}")

    return {
        'protocol': parsed_url.scheme,
        'path': parsed_url.path,
        'query': decoded_query,
        'data': decoded_data
    }


def parse_content_type(header_str):
    lines = [line.strip() for line in header_str.split('\n') if line.strip()]

    content_type = None
    accept = None

    for line in lines:
        if ':' not in line:
            continue  # 跳过非头部行（例如 GET、路径、HTTP版本）

        key, value = line.split(':', 1)
        key = key.strip()
        value = value.strip()

        # 先检查 Content-Type
        if key == 'Content-Type' and content_type is None:
            content_type = value.split(',')[0].strip()

        # 如果还没找到 Content-Type，再检查 Accept
        elif key == 'Accept' and accept is None and content_type is None:
            accept = value.split(',')[0].strip()  # 取第一个值

    # 按优先级返回结果
    if content_type:
        return content_type
    elif accept:
        return accept
    else:
        return None

def parse_request_headers(header_str):
    """解析请求头字段"""
    method = None

    if header_str.strip()[:6] == 'DELETE':
        method = 'delete'

    elif header_str.strip()[:3] == 'GET':
        method = 'get'

    elif header_str.strip()[:3] == 'PUT':
        method = 'put'

    elif header_str.strip()[:4] == 'POST':
        method = 'post'

    headers = parse_content_type(header_str)

    return {
        'method': method,
        'headers': headers
    }


# ---------------------- 数据处理逻辑 ----------------------
CONTENT_TYPE_MAP = {
    "application/json": 1,
    "application/x-www-form-urlencoded": 2,
    "multipart/form-data": 3,
    "text/plain": 4,
    "text/xml": 5,
    "tcbs一代客户端": 8,
    "application/x-amf": 9
}


def read_input_excel(input_file):
    """读取输入Excel文件"""
    wb = load_workbook(input_file)
    ws = wb.active
    data = {}

    for row in ws.iter_rows(min_row=2):
        if len(row) >=5 and row[0].value and row[1].value:
            interface_num = str(row[0].value).strip()
            interface_name = str(row[1].value).strip()
            curl_command = str(row[2].value).strip()
            request_headers = str(row[3].value).strip()
            response = str(row[4].value).strip() if row[4].value else ''
            data[interface_name] = (
                interface_num,
                curl_command,
                request_headers,
                response
            )
    return data


def generate_interface_data(interface_name, parsed_data, interface_num, headers):
    """生成接口文档数据行"""
    # 设置header字段
    if interface_name == "总行管理员登录":
        header_value = None
    else:
        header_value = 'Authorization=123'

    # 确定编码类型
    content_type = parse_request_headers(header_str=headers)['headers']
    if content_type is not None:
        encoding_type = CONTENT_TYPE_MAP[content_type]
    else:
        encoding_type = 7

    # 确定请求方法
    method = parse_request_headers(header_str=headers)['method']

    return {
        '接口编号': interface_num,
        '项目名称': '请填写',
        '接口名称': interface_name,
        '接口协议(http、https)': parsed_data['protocol'],
        '接口路径': parsed_data['path'],
        '接口请求方法(get、post、put、delete)': method,
        '请求体编码类型(1:"application/json", 2:"application/x-www-form-urlencoded", 3:"multipart/form-data", 4:"text/plain", 5:"text/xml", 7:"none", 8:"TCBS一代客户端", 9:"application/x-amf")': encoding_type,  # 修改为数字
        'header': header_value,
        'query-params(http请求参数，以@@分割拼接在url后)': parsed_data['query'],
        'body(http请求体)': parsed_data['data']
    }


def process_body_data(parsed_data):
    """处理body数据，根据内容类型转换为键值对格式"""
    content_type = parsed_data.get('content_type', '')
    data = parsed_data.get('data', '')

    # 判断是否为JSON格式
    if 'application/json' in content_type.lower() and data:
        try:
            json_data = json.loads(data)
            items = []
            for key, value in json_data.items():
                # 处理不同数据类型
                if isinstance(value, (list, dict)):
                    items.append(f"{key}={json.dumps(value, ensure_ascii=False)}")
                else:
                    items.append(f"{key}={value}")
            return '@@'.join(items)
        except (json.JSONDecodeError, AttributeError):
            # 解析失败时保留原始数据
            return data
    else:
        # 非JSON数据直接替换&符号
        return data.replace('&', '@@') if data else ''


def generate_testcase_rows(interface_name, parsed_data, response, case_num):
    """生成测试用例的多行数据"""
    # 设置header替换项
    if interface_name == "总行管理员登录":
        header_value = None
    else:
        header_value = 'Authorization=Bearer #{总行管理员登录_RECMSG}_token#'

    # 处理query-params字段（替换&为@@）
    processed_body = process_body_data(parsed_data)

    # 使用有序字典确保字段顺序
    base_data = OrderedDict([
        ('用例状态(0=不执行@@1=发送+接收@@2=仅接收)', 1),
        ('用例编号', f"{interface_name}_{case_num:02d}"),
        ('中间件类型', 'HTTP'),
        ('其他配置项', ''),
        ('接口名称', interface_name),
        ('用例名称', f"{interface_name}_{case_num:02d}"),
        ('前置动作', ''),
        ('header(http请求头参数替换项)', header_value),
        # ('接口路径(接口路径参数替换项)', parsed_data['path']),
        ('接口路径(接口路径参数替换项)', ''),
        ('query-params(http请求参数替换项)', parsed_data['query']),
        ('body(http请求体参数替换项)', processed_body),
        ('校验动作', ''),
        ('校验项', ''),
        ('断言方式', ''),
        ('预期值', ''),
        ('后置动作', '')
    ])

    assertions = []
    if response:
        # 断言逻辑待处理
        if response == 'PK':
            assertion_configs = [
                {
                    'field': response.strip().content[:2],
                    'assertion_type': '等于',
                    'value': 'PK'
                },
                {
                    'field': response.status_code,
                    'assertion_type': '等于',
                    'value': 200
                },
                {
                    'field': response.headers['Content-Type'],
                    'assertion_type': '包含',
                    'value': 'application/vod.ms-excel'
                },
                # 可继续扩展其他断言内容
            ]

            for config in assertion_configs:
                assertions.append({
                    '校验动作': '',
                    '校验项': config['field'],
                    '断言方式': config['assertion_type'],
                    '预期值': config['value']
                })

        else:
            processed_responses = response_process.load_json(response)
            for processed_response in processed_responses:
                for key, value in processed_response.items():
                    assertions.append({
                        '校验动作': '',
                        '校验项': key,
                        '断言方式': "等于",
                        '预期值': value
                    })

    else:
        assertions.append({
            '校验动作': '',
            '校验项': "status_code",
            '断言方式': "等于",
            '预期值': 200
        })

    # 生成多行数据
    rows = []
    for assertion in assertions:
        row = base_data.copy()
        row.update(assertion)
        rows.append(row)

    # 需要合并的列范围（前11列：A-K）
    merge_columns = list(range(11))  # 对应OrderedDict前11个字段
    return rows, len(rows), merge_columns


def main(input_file="./cUrl命令.xlsx"):
    # 读取输入文件
    try:
        curl_data = read_input_excel(input_file)
    except FileNotFoundError:
        print(f"错误：输入文件 {input_file} 不存在")
        return
    except Exception as e:
        print(f"读取Excel失败：{str(e)}")
        return

    # 创建两个独立的工作簿
    interface_wb = Workbook()
    testcase_wb = Workbook()

    # 删除默认创建的Sheet
    for wb in [interface_wb, testcase_wb]:
        if 'Sheet' in wb.sheetnames:
            del wb['Sheet']

    # ==================== 样式配置 ====================
    # 公共样式配置
    header_font = Font(name='宋体', size=12, bold=True)
    header_fill = PatternFill(start_color='B4C6E7', end_color='B4C6E7', fill_type='solid')
    header_alignment = Alignment(wrap_text=True, horizontal='center', vertical='center', text_rotation=0)
    column_width = 40

    # ==================== 接口文档表 ====================
    interface_sheet = interface_wb.create_sheet("接口文档")
    interface_headers = [
        "接口编号",
        "项目名称",
        "接口名称",
        "接口协议(http、https)",
        "接口路径",
        "接口请求方法(get、post、put、delete)",
        '请求体编码类型(1:"application/json", 2:"application/x-www-form-urlencoded", 3:"multipart/form-data", 4:"text/plain", 5:"text/xml", 7:"none", 8:"TCBS一代客户端", 9:"application/x-amf")',
        "header(http请求头)",
        "query-params(http请求参数，以@@分割拼接在url后)",
        "body(http请求体)"
    ]
    interface_sheet.append(interface_headers)

    # 设置列宽和表头样式
    for col in range(1, len(interface_headers) + 1):
        col_letter = get_column_letter(col)
        interface_sheet.column_dimensions[col_letter].width = column_width

    for cell in interface_sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    # ==================== 测试用例表 ====================
    testcase_sheet = testcase_wb.create_sheet("测试用例")
    testcase_headers = [
        "用例状态(0=不执行@@1=发送+接收@@2=仅接收)", "用例编号", "中间件类型", "其他配置项", "接口名称",
        "用例名称", "前置动作", "header(http请求头参数替换项)", "接口路径(接口路径参数替换项)",
        "query-params(http请求参数替换项)",
        "body(http请求体参数替换项)", "校验动作", "校验项", "断言方式", "预期值", "后置动作"
    ]
    testcase_sheet.append(testcase_headers)

    # 设置列宽和表头样式
    for col in range(1, len(testcase_headers) + 1):
        col_letter = get_column_letter(col)
        testcase_sheet.column_dimensions[col_letter].width = column_width

    for cell in testcase_sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    # 处理每个接口
    current_row = 2  # 数据起始行
    for interface_name, (interface_num, curl_command, headers, response) in curl_data.items():
        try:
            parsed = parse_curl(curl_command)

            # 生成并写入接口文档
            interface_row = generate_interface_data(interface_name, parsed, interface_num, headers)
            interface_sheet.append(list(interface_row.values()))

            # 生成测试用例数据
            case_num = 1  # 每个接口默认生成一个用例（含多个断言）
            rows, row_count, merge_cols = generate_testcase_rows(interface_name, parsed, response, case_num)

            # 写入测试用例并合并单元格
            start_row = current_row
            for row in rows:
                testcase_sheet.append(list(row.values()))
                current_row += 1

                # 合并前11列单元格（A-K）
                end_row = start_row + row_count - 1
                for col_idx in merge_cols:
                    col_letter = get_column_letter(col_idx + 1)
                    testcase_sheet.merge_cells(f'{col_letter}{start_row}:{col_letter}{end_row}')
                    # 设置垂直居中
                    for row in range(start_row, end_row + 1):
                        testcase_sheet[f'{col_letter}{row}'].alignment = Alignment(vertical='center')

            print(f"成功处理接口：{interface_name}")

        except Exception as e:
            print(f"处理接口 [{interface_name}] 时出错：{str(e)}")
            continue

    # 保存文件
    try:
        interface_wb.save("接口文档.xlsx")
        testcase_wb.save("测试用例.xlsx")
        print("成功生成文件：\n- 接口文档.xlsx\n- 测试用例.xlsx")
    except PermissionError:
        print("错误：请关闭正在使用的Excel文件后重试")


if __name__ == "__main__":
    main()
