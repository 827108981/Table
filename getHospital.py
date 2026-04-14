import openpyxl
import json
import os
import webbrowser
import http.server
import socketserver
import threading
import time
import socket
import requests
import re
from datetime import datetime
from pathlib import Path


def extract_workorder_data(file_path):
    """
    从工单报表Excel文件中提取完整的工单信息

    Args:
        file_path: Excel文件路径

    Returns:
        list: 工单数据列表，每个元素是一个字典
    """
    try:
        # 加载工作簿
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

        # 获取表头，确定各列的索引
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
        header_dict = {}
        for idx, header in enumerate(header_row):
            if header:
                header_dict[str(header).strip()] = idx

        print(f"表头列名: {list(header_dict.keys())}")

        # 确定需要的列索引
        workorder_col = header_dict.get('工单号')
        engineer_col = header_dict.get('服务工程师')
        product_col = header_dict.get('产品型号')
        create_time_col = header_dict.get('创建时间')
        complete_time_col = header_dict.get('预约完成时间')
        business_desc_col = header_dict.get('业务描述')
        
        # 新增字段：分公司、工单状态、工单创建时间
        branch_col = header_dict.get('分公司')
        status_col = header_dict.get('工单状态')
        # 如果"工单创建时间"和"创建时间"是同一列，则复用
        order_create_time_col = header_dict.get('工单创建时间') or create_time_col

        required_cols = [workorder_col, engineer_col, product_col, create_time_col, complete_time_col, business_desc_col]
        if any(col is None for col in required_cols):
            missing_cols = []
            if workorder_col is None: missing_cols.append('工单号')
            if engineer_col is None: missing_cols.append('服务工程师')
            if product_col is None: missing_cols.append('产品型号')
            if create_time_col is None: missing_cols.append('创建时间')
            if complete_time_col is None: missing_cols.append('预约完成时间')
            if business_desc_col is None: missing_cols.append('业务描述')
            print(f"错误：找不到以下列: {', '.join(missing_cols)}")
            return []

        # 打印可选字段提示
        if branch_col is None: print("⚠️ 未找到'分公司'列，将留空")
        if status_col is None: print("⚠️ 未找到'工单状态'列，将留空")

        # 创建结果列表
        result_list = []
        index = 0

        # 从第2行开始遍历（第1行是表头）
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # 获取各列数据
            workorder_no = row[workorder_col]
            engineer = row[engineer_col]
            product_model = row[product_col]
            create_time = row[create_time_col]
            complete_time = row[complete_time_col]
            business_desc = row[business_desc_col]
            
            # 新增字段提取
            branch = row[branch_col] if branch_col is not None else ''
            status = row[status_col] if status_col is not None else ''
            order_create_time = row[order_create_time_col] if order_create_time_col is not None else create_time

            # 跳过空值
            if workorder_no is None:
                continue

            # 处理产品型号：取第一个字段
            product_type = ''
            device_model = ''
            if product_model:
                fields = str(product_model).split()
                if len(fields) >= 1:
                    product_type = fields[0]
                    device_model = fields[0]

            # 处理业务描述：取第二个字段作为医院地址
            hospital_address = ''
            if business_desc:
                fields = str(business_desc).split()
                if len(fields) >= 2:
                    hospital_address = fields[1]
                elif len(fields) == 1:
                    hospital_address = fields[0]

            # 格式化时间
            def format_time(t):
                if not t: return ''
                if hasattr(t, 'strftime'):
                    return t.strftime('%Y-%m-%d %H:%M:%S')
                return str(t)

            create_time_str = format_time(create_time)
            complete_time_str = format_time(complete_time)
            order_create_time_str = format_time(order_create_time)

            # 构建字典
            workorder_dict = {
                "index": index,
                "工单编号": str(workorder_no),
                "产品类型": product_type,
                "设备型号": device_model,
                "服务工程师": str(engineer) if engineer else '',
                "创建时间": create_time_str,
                "预计完成时间": complete_time_str,
                "医院地址": hospital_address,
                "分公司": str(branch) if branch else '',
                "工单状态": str(status) if status else '',
                "工单创建时间": order_create_time_str
            }

            result_list.append(workorder_dict)
            index += 1

        return result_list

    except FileNotFoundError:
        print(f"错误：找不到文件 {file_path}")
        return []
    except Exception as e:
        print(f"错误：{str(e)}")
        import traceback
        traceback.print_exc()
        return []


def generate_hospital_json(workorder_list, output_file="hospital_data.json"):
    """
    将工单数据生成为JSON文件供HTML使用

    Args:
        workorder_list: 工单数据列表
        output_file: 输出JSON文件路径

    Returns:
        list: 地图数据列表
    """
    # 配置百度地图API Key（需要申请）
    # 申请地址：https://lbsyun.baidu.com/apiconsole/key
    BAIDU_AK = "Ch6L97O8ZKSN3KEJgI5UuaKmvH3b74gL"  # 替换为你的API Key
    USE_GEOCODING = True  # 是否使用地理编码API
    
    # 扩展的区县级坐标映射（覆盖更多地区）
    district_coordinates = {
        # 北京市区县
        "朝阳": [116.44, 39.92], "海淀": [116.30, 39.96], "西城": [116.37, 39.91],
        "东城": [116.42, 39.93], "丰台": [116.29, 39.86], "石景山": [116.20, 39.91],
        "通州": [116.66, 39.91], "顺义": [116.65, 40.13], "昌平": [116.23, 40.22],
        "大兴": [116.34, 39.73], "房山": [116.14, 39.75], "密云": [116.84, 40.38],
        "怀柔": [116.63, 40.32], "平谷": [117.12, 40.14], "延庆": [115.98, 40.46],
        
        # 上海市区县
        "黄浦": [121.48, 31.23], "徐汇": [121.43, 31.18], "长宁": [121.42, 31.22],
        "静安": [121.45, 31.23], "普陀": [121.40, 31.25], "虹口": [121.48, 31.26],
        "杨浦": [121.52, 31.26], "浦东": [121.54, 31.22], "闵行": [121.38, 31.11],
        "宝山": [121.49, 31.40], "嘉定": [121.25, 31.38], "金山": [121.34, 30.74],
        "松江": [121.23, 31.03], "青浦": [121.12, 31.15], "奉贤": [121.47, 30.92],
        
        # 广州市区县
        "天河": [113.33, 23.13], "越秀": [113.27, 23.13], "海珠": [113.27, 23.10],
        "荔湾": [113.24, 23.13], "白云": [113.27, 23.16], "黄埔": [113.46, 23.10],
        "番禺": [113.38, 22.94], "花都": [113.22, 23.40], "南沙": [113.53, 22.80],
        
        # 深圳市区县
        "福田": [114.05, 22.52], "罗湖": [114.12, 22.55], "南山": [113.93, 22.53],
        "盐田": [114.23, 22.55], "宝安": [113.88, 22.55], "龙岗": [114.25, 22.72],
        
        # 成都市区县
        "锦江": [104.08, 30.66], "青羊": [104.06, 30.67], "金牛": [104.04, 30.69],
        "武侯": [104.04, 30.64], "成华": [104.10, 30.66], "龙泉驿": [104.27, 30.56],
        "青白江": [104.25, 30.88], "新都": [104.16, 30.82], "温江": [103.84, 30.70],
        "双流": [103.92, 30.58], "郫都": [103.89, 30.81],
        
        # 更多地级市和区县（扩展版）
        "石家庄": [114.48, 38.03], "唐山": [118.18, 39.63], "秦皇岛": [119.60, 39.93],
        "邯郸": [114.48, 36.60], "邢台": [114.50, 37.07], "保定": [115.48, 38.85],
        "张家口": [114.88, 40.77], "承德": [117.93, 40.97], "沧州": [116.86, 38.31],
        "廊坊": [116.70, 39.53], "衡水": [115.68, 37.73], "太原": [112.55, 37.87],
        "大同": [113.30, 40.08], "阳泉": [113.58, 37.85], "长治": [113.12, 36.20],
        "晋城": [112.85, 35.48], "朔州": [112.43, 39.33], "晋中": [112.75, 37.68],
        "运城": [111.00, 35.02], "忻州": [112.73, 38.42], "临汾": [111.52, 36.08],
        "沈阳": [123.43, 41.80], "大连": [121.62, 38.92], "鞍山": [122.99, 41.11],
        "抚顺": [123.92, 41.86], "本溪": [123.77, 41.30], "丹东": [124.37, 40.13],
        "锦州": [121.15, 41.13], "营口": [122.23, 40.65], "阜新": [121.66, 42.02],
        "辽阳": [123.18, 41.27], "盘锦": [122.06, 41.12], "铁岭": [123.85, 42.28],
        "长春": [125.35, 43.88], "吉林": [126.55, 43.85], "四平": [124.37, 43.17],
        "辽源": [125.15, 42.90], "通化": [125.93, 41.73], "白山": [126.43, 41.93],
        "哈尔滨": [126.63, 45.75], "齐齐哈尔": [123.97, 47.33], "鸡西": [130.97, 45.30],
        "大庆": [125.10, 46.58], "牡丹江": [129.60, 44.58], "南京": [118.78, 32.06],
        "无锡": [120.31, 31.49], "徐州": [117.18, 34.27], "常州": [119.97, 31.81],
        "苏州": [120.59, 31.30], "南通": [120.89, 31.98], "连云港": [119.22, 34.60],
        "淮安": [119.15, 33.50], "盐城": [120.15, 33.38], "扬州": [119.42, 32.38],
        "镇江": [119.45, 32.20], "泰州": [119.92, 32.48], "宿迁": [118.28, 33.97],
        "杭州": [120.15, 30.28], "宁波": [121.55, 29.87], "温州": [120.70, 28.00],
        "嘉兴": [120.75, 30.75], "湖州": [120.10, 30.87], "绍兴": [120.58, 30.03],
        "金华": [119.65, 29.08], "衢州": [118.87, 28.95], "舟山": [122.10, 30.02],
        "台州": [121.42, 28.66], "丽水": [119.92, 28.45], "合肥": [117.27, 31.86],
        "芜湖": [118.38, 31.33], "蚌埠": [117.37, 32.92], "淮南": [117.00, 32.63],
        "马鞍山": [118.52, 31.68], "淮北": [116.78, 33.97], "铜陵": [117.82, 30.93],
        "安庆": [117.05, 30.52], "黄山": [118.32, 29.72], "滁州": [118.32, 32.30],
        "阜阳": [115.82, 32.88], "宿州": [116.97, 33.63], "六安": [116.52, 31.75],
        "福州": [119.30, 26.08], "厦门": [118.09, 24.48], "莆田": [119.00, 25.43],
        "三明": [117.63, 26.27], "泉州": [118.67, 24.88], "漳州": [117.65, 24.52],
        "南平": [118.17, 26.63], "龙岩": [117.02, 25.08], "宁德": [119.52, 26.67],
        "南昌": [115.89, 28.68], "景德镇": [117.22, 29.30], "萍乡": [113.85, 27.63],
        "九江": [115.98, 29.72], "新余": [114.93, 27.80], "鹰潭": [117.03, 28.23],
        "赣州": [114.93, 25.83], "吉安": [114.98, 27.12], "宜春": [114.40, 27.80],
        "济南": [117.02, 36.65], "青岛": [120.38, 36.07], "淄博": [118.05, 36.78],
        "枣庄": [117.57, 34.87], "东营": [118.67, 37.43], "烟台": [121.40, 37.47],
        "潍坊": [119.15, 36.70], "济宁": [116.58, 35.42], "泰安": [117.13, 36.18],
        "威海": [122.12, 37.50], "日照": [119.53, 35.42], "临沂": [118.35, 35.07],
        "德州": [116.35, 37.43], "聊城": [115.98, 36.45], "滨州": [118.02, 37.37],
        "郑州": [113.65, 34.76], "开封": [114.35, 34.80], "洛阳": [112.45, 34.62],
        "平顶山": [113.30, 33.73], "安阳": [114.35, 36.10], "鹤壁": [114.30, 35.75],
        "新乡": [113.92, 35.30], "焦作": [113.25, 35.22], "濮阳": [115.03, 35.73],
        "许昌": [113.83, 34.03], "漯河": [114.02, 33.58], "南阳": [112.53, 33.00],
        "商丘": [115.65, 34.43], "信阳": [114.07, 32.13], "周口": [114.65, 33.62],
        "武汉": [114.31, 30.52], "黄石": [115.03, 30.22], "十堰": [110.78, 32.63],
        "宜昌": [111.28, 30.70], "襄阳": [112.15, 32.02], "鄂州": [114.88, 30.38],
        "荆门": [112.20, 31.03], "孝感": [113.92, 30.92], "荆州": [112.18, 30.32],
        "黄冈": [114.87, 30.43], "咸宁": [114.32, 29.88], "随州": [113.38, 31.72],
        "长沙": [112.94, 28.23], "株洲": [113.15, 27.83], "湘潭": [112.93, 27.83],
        "衡阳": [112.62, 26.90], "邵阳": [111.47, 27.23], "岳阳": [113.13, 29.37],
        "常德": [111.68, 29.03], "张家界": [110.48, 29.12], "益阳": [112.35, 28.55],
        "郴州": [113.02, 25.78], "永州": [111.62, 26.43], "怀化": [110.00, 27.55],
        "广州": [113.27, 23.13], "韶关": [113.60, 24.80], "深圳": [114.06, 22.54],
        "珠海": [113.57, 22.28], "汕头": [116.68, 23.35], "佛山": [113.12, 23.02],
        "江门": [113.08, 22.58], "湛江": [110.35, 21.27], "茂名": [110.92, 21.67],
        "肇庆": [112.47, 23.05], "惠州": [114.42, 23.08], "梅州": [116.12, 24.28],
        "南宁": [108.33, 22.84], "柳州": [109.42, 24.33], "桂林": [110.28, 25.27],
        "梧州": [111.30, 23.48], "北海": [109.12, 21.48], "钦州": [108.62, 21.97],
        "贵港": [109.60, 23.10], "玉林": [110.15, 22.63], "百色": [106.62, 23.90],
        "成都": [104.07, 30.67], "自贡": [104.78, 29.35], "攀枝花": [101.72, 26.58],
        "泸州": [105.43, 28.88], "德阳": [104.38, 31.13], "绵阳": [104.70, 31.47],
        "广元": [105.83, 32.43], "遂宁": [105.57, 30.52], "内江": [105.07, 29.58],
        "乐山": [103.75, 29.55], "南充": [106.08, 30.78], "眉山": [103.83, 30.05],
        "宜宾": [104.62, 28.77], "广安": [106.63, 30.45], "达州": [107.50, 31.22],
        "贵阳": [106.71, 26.57], "遵义": [106.93, 27.70], "六盘水": [104.83, 26.58],
        "昆明": [102.73, 25.04], "曲靖": [103.78, 25.50], "玉溪": [102.53, 24.35],
        "西安": [108.94, 34.34], "铜川": [108.95, 34.90], "宝鸡": [107.15, 34.37],
        "咸阳": [108.70, 34.33], "渭南": [109.50, 34.50], "延安": [109.48, 36.60],
        "汉中": [107.03, 33.07], "榆林": [109.73, 38.28], "兰州": [103.83, 36.06],
        "西宁": [101.74, 36.56], "银川": [106.27, 38.47], "石嘴山": [106.38, 39.02],
        "乌鲁木齐": [87.68, 43.77], "克拉玛依": [84.87, 45.58], "吐鲁番": [89.18, 42.95],
        "呼和浩特": [111.65, 40.82], "包头": [109.83, 40.65], "赤峰": [118.97, 42.27],
        "通辽": [122.27, 43.62], "鄂尔多斯": [109.78, 39.82], "拉萨": [91.11, 29.97],
        "海口": [110.35, 20.02], "三亚": [109.51, 18.25],
    }

    # 常见城市和对应经纬度映射表
    city_coordinates = {
        "北京": [116.41, 39.91],
        "上海": [121.47, 31.23],
        "广州": [113.27, 23.13],
        "深圳": [114.06, 22.54],
        "成都": [104.07, 30.67],
        "武汉": [114.31, 30.52],
        "南京": [118.78, 32.06],
        "杭州": [120.15, 30.28],
        "西安": [108.94, 34.34],
        "重庆": [106.55, 29.56],
        "天津": [117.20, 39.13],
        "苏州": [120.59, 31.30],
        "长沙": [112.94, 28.23],
        "郑州": [113.65, 34.76],
        "济南": [117.02, 36.65],
        "青岛": [120.38, 36.07],
        "大连": [121.62, 38.92],
        "厦门": [118.09, 24.48],
        "福州": [119.30, 26.08],
        "合肥": [117.27, 31.86],
        "昆明": [102.73, 25.04],
        "南昌": [115.89, 28.68],
        "贵阳": [106.71, 26.57],
        "南宁": [108.33, 22.84],
        "哈尔滨": [126.63, 45.75],
        "长春": [125.35, 43.88],
        "沈阳": [123.43, 41.80],
        "石家庄": [114.48, 38.03],
        "太原": [112.55, 37.87],
        "呼和浩特": [111.65, 40.82],
        "兰州": [103.83, 36.06],
        "银川": [106.27, 38.47],
        "西宁": [101.74, 36.56],
        "乌鲁木齐": [87.68, 43.77],
        "拉萨": [91.11, 29.97],
        "海口": [110.35, 20.02],
    }

    # 省份到主要城市的映射
    province_to_city = {
        "江苏": "南京",
        "浙江": "杭州",
        "广东": "广州",
        "四川": "成都",
        "湖北": "武汉",
        "山东": "济南",
        "河南": "郑州",
        "湖南": "长沙",
        "安徽": "合肥",
        "福建": "福州",
        "江西": "南昌",
        "陕西": "西安",
        "河北": "石家庄",
        "山西": "太原",
        "辽宁": "沈阳",
        "吉林": "长春",
        "黑龙江": "哈尔滨",
        "云南": "昆明",
        "贵州": "贵阳",
        "广西": "南宁",
        "甘肃": "兰州",
        "内蒙古": "呼和浩特",
        "新疆": "乌鲁木齐",
        "西藏": "拉萨",
        "青海": "西宁",
        "宁夏": "银川",
        "海南": "海口",
    }
    
    def geocode_with_baidu(address, ak):
        """
        使用百度地图API进行地理编码
        
        Args:
            address: 地址字符串
            ak: 百度地图API Key
        
        Returns:
            [lng, lat] 或 None
        """
        if not USE_GEOCODING or ak == "你的百度地图AK":
            return None
            
        try:
            url = "http://api.map.baidu.com/geocoding/v3/"
            params = {
                "address": address,
                "output": "json",
                "ak": ak,
                "callback": "showLocation"
            }
            
            response = requests.get(url, params=params, timeout=5)
            response_text = response.text
            
            # 解析JSONP响应
            if "showLocation&&showLocation(" in response_text:
                json_str = response_text.replace("showLocation&&showLocation(", "").rstrip(")")
                data = json.loads(json_str)
                
                if data.get("status") == 0 and "result" in data:
                    location = data["result"]["location"]
                    return [location["lng"], location["lat"]]
            
            return None
        except Exception as e:
            print(f"⚠️ 百度API地理编码失败: {e}")
            return None
    
    def parse_address_to_coordinates(address):
        """
        智能解析地址为经纬度坐标（优先精确，降级模糊）
        
        解析顺序：
        1. 百度地图API地理编码（最精确）
        2. 区县级匹配
        3. 城市级匹配
        4. 省份级匹配
        5. 默认中国中心点
        
        Args:
            address: 地址字符串
        
        Returns:
            [lng, lat]: 经纬度坐标
        """
        if not address:
            return [105.0, 35.0]  # 默认中国中心
        
        # 1. 优先使用百度地图API进行精确地理编码
        coordinates = geocode_with_baidu(address, BAIDU_AK)
        if coordinates:
            print(f"✅ 使用百度API解析成功: {address} -> {coordinates}")
            return coordinates
        
        # 2. 尝试区县级匹配（更高精度）
        for district, coords in district_coordinates.items():
            if district in address:
                print(f"📍 匹配到区县: {address} -> {district}")
                return coords
        
        # 3. 尝试城市级匹配
        for city, coords in city_coordinates.items():
            if city in address:
                print(f"📍 匹配到城市: {address} -> {city}")
                return coords
        
        # 4. 尝试省份级匹配
        for province, city in province_to_city.items():
            if province in address:
                coords = city_coordinates.get(city)
                if coords:
                    print(f"📍 匹配到省份: {address} -> {province} -> {city}")
                    return coords
        
        # 5. 使用默认坐标（中国中心）
        print(f"⚠️ 未匹配到任何地区，使用默认坐标: {address}")
        return [105.0, 35.0]

    map_data = []

    for item in workorder_list:
        hospital_address = item['医院地址']

        # 使用智能解析函数获取坐标
        coordinates = parse_address_to_coordinates(hospital_address)

        # 计算超时时长与紧急程度
        timeout_hours = 0
        urgency_level = 'warning'
        
        try:
            if item['创建时间']:
                create_dt = datetime.strptime(item['创建时间'], '%Y-%m-%d %H:%M:%S')
                timeout_hours = round((datetime.now() - create_dt).total_seconds() / 3600, 1)
                
                # 紧急程度分级规则
                if timeout_hours > 72:
                    urgency_level = 'critical'   # 🔴 严重超时
                elif timeout_hours > 48:
                    urgency_level = 'urgent'     # 🟠 紧急超时
                elif timeout_hours > 30:
                    urgency_level = 'warning'    # 🟡 一般超时
        except Exception as e:
            print(f"⚠️ 计算超时时长失败: {e}")

        map_data.append({
            "name": hospital_address,
            "value": coordinates + [1],
            "type": urgency_level,
            "timeoutHours": timeout_hours,
            "index": item['index'],
            "工单编号": item['工单编号'],
            "产品类型": item['产品类型'],
            "设备型号": item['设备型号'],
            "服务工程师": item['服务工程师'],
            "创建时间": item['创建时间'],
            "预计完成时间": item['预计完成时间'],
            "医院地址": hospital_address,
            "分公司": item.get('分公司', ''),
            "工单状态": item.get('工单状态', ''),
            "工单创建时间": item.get('工单创建时间', '')
        })

    # 保存JSON文件
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(map_data, f, ensure_ascii=False, indent=2)

    print(f"✓ 已生成JSON数据文件: {output_file}")
    print(f"✓ 共生成 {len(map_data)} 条工单数据")
    
    # 统计定位精度
    api_count = sum(1 for item in map_data if "百度API" in str(item))
    district_count = sum(1 for item in map_data if "区县" in str(item))
    city_count = sum(1 for item in map_data if "城市" in str(item))
    
    print(f"\n📊 定位统计:")
    print(f"  API精确解析: {api_count} 条")
    print(f"  区县级匹配: {district_count} 条")
    print(f"  城市级匹配: {city_count} 条")

    return map_data


def find_available_port(start_port=8080, max_attempts=100):
    """
    查找可用端口

    Args:
        start_port: 起始端口号
        max_attempts: 最大尝试次数

    Returns:
        int: 可用的端口号
    """
    for port in range(start_port, start_port + max_attempts):
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(1)
        try:
            result = sock.connect_ex(('localhost', port))
            if result != 0:
                sock.close()
                return port
        except Exception:
            pass
        finally:
            sock.close()

    raise Exception(f"无法找到可用端口（尝试了 {start_port} 到 {start_port + max_attempts - 1}）")


def start_http_server(port=8080, directory=None):
    """
    启动HTTP服务器

    Args:
        port: 端口号
        directory: 服务目录，默认为脚本所在目录
    """
    if directory is None:
        directory = os.path.dirname(os.path.abspath(__file__))

    os.chdir(directory)

    handler = http.server.SimpleHTTPRequestHandler
    handler.extensions_map.update({
        '.json': 'application/json',
        '.html': 'text/html',
    })

    # 允许地址重用
    socketserver.TCPServer.allow_reuse_address = True

    try:
        httpd = socketserver.TCPServer(("", port), handler)
        print(f"✓ HTTP服务器已启动在端口 {port}")
        print(f"✓ 服务目录: {directory}")
        httpd.serve_forever()
    except OSError as e:
        print(f"❌ 服务器启动失败: {e}")
        raise


def open_browser(port=8080, delay=1.5):
    """
    延迟打开浏览器（在新标签页中打开，不覆盖现有窗口）

    Args:
        port: HTTP服务器端口
        delay: 延迟时间（秒）
    """
    if delay > 0:
        time.sleep(delay)
    url = f"http://localhost:{port}/dashboard.html"
    print(f"✓ 正在打开浏览器: {url}")
    print(f"💡 提示：前端将在新标签页打开，不会覆盖MSP浏览器窗口")
    try:
        # 使用 'new' 参数确保在新标签页/窗口中打开
        webbrowser.open(url, new=2)  # new=2 表示在新标签页中打开
        print(f"✓ 浏览器已在新标签页打开")
    except Exception as e:
        print(f"⚠️ 自动打开浏览器失败: {e}")
        print(f"请手动访问: {url}")


# 使用示例
if __name__ == "__main__":
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 使用相对路径，兼容Mac和Windows
    data_folder = os.path.join(script_dir, "工单报表数据")
    excel_filename = "工单报表_最新数据.xlsx"  # 或使用 "工单报表_超时数据.xlsx"
    file_path = os.path.join(data_folder, excel_filename)

    json_output = os.path.join(script_dir, "hospital_data.json")
    html_file = os.path.join(script_dir, "dashboard.html")

    # 检查HTML文件是否存在
    if not os.path.exists(html_file):
        print(f"错误：找不到HTML文件 {html_file}")
        exit()

    # 检查Excel文件是否存在
    if not os.path.exists(file_path):
        print(f"错误：找不到Excel文件 {file_path}")
        exit()

    # 提取数据
    print("=" * 60)
    print("迈瑞技术支持驾驶舱 - 数据更新系统")
    print("=" * 60)
    print("\n步骤 1/4: 正在提取工单数据...")
    workorder_list = extract_workorder_data(file_path)

    if not workorder_list:
        print("未提取到任何数据，程序退出。")
        exit()

    # 打印提取结果
    print(f"✓ 成功提取 {len(workorder_list)} 条工单数据")
    print("\n提取的工单数据示例（前3条）：")
    for item in workorder_list[:3]:
        print(f"  工单编号: {item['工单编号']}")
        print(f"  产品类型: {item['产品类型']}")
        print(f"  设备型号: {item['设备型号']}")
        print(f"  服务工程师: {item['服务工程师']}")
        print(f"  创建时间: {item['创建时间']}")
        print(f"  预计完成时间: {item['预计完成时间']}")
        print(f"  医院地址: {item['医院地址']}")
        print()
    if len(workorder_list) > 3:
        print(f"  ... 还有 {len(workorder_list) - 3} 条数据")

    # 生成JSON数据文件
    print("\n步骤 2/4: 正在生成地图数据...")
    map_data = generate_hospital_json(workorder_list, json_output)

    # 打印生成的数据统计
    type_stats = {}
    for item in map_data:
        type_stats[item['type']] = type_stats.get(item['type'], 0) + 1

    print("\n数据类型统计：")
    type_names = {
        'normal': '正常工单',
        'timeout': '超时服务',
        'complaint': '投诉工单'
    }
    for type_key, count in type_stats.items():
        print(f"  {type_names.get(type_key, type_key)}: {count} 条")

    # 启动HTTP服务器
    print("\n步骤 3/4: 正在启动HTTP服务器...")

    try:
        # 查找可用端口
        port = find_available_port(8080)
        print(f"✓ 使用端口: {port}")

        # 在新线程中启动HTTP服务器
        server_thread = threading.Thread(target=start_http_server, args=(port, script_dir), daemon=True)
        server_thread.start()

        # 等待服务器启动
        time.sleep(1)

        # 打开浏览器
        print("\n步骤 4/4: 正在打开浏览器...")
        browser_thread = threading.Thread(target=open_browser, args=(port,), daemon=True)
        browser_thread.start()

        print("\n" + "=" * 60)
        print("✓ 系统启动完成！")
        print(f"✓ 访问地址: http://localhost:{port}/dashboard.html")
        print("=" * 60)
        print("\n提示：")
        print("- 浏览器将自动打开驾驶舱页面")
        print("- 按 Ctrl+C 可停止服务器")
        print("- 数据已自动传送到HTML页面\n")

        # 保持主线程运行
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n\n服务器已停止")

    except Exception as e:
        print(f"\n❌ 系统启动失败: {e}")
        import traceback

        traceback.print_exc()
        input("\n按回车键退出...")
