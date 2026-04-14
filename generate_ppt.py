"""
将 Markdown 汇报文档转换为 PPT 演示文稿
依赖：pip install python-pptx
"""
import re
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# 颜色配置
MINDRAY_BLUE = RGBColor(0, 176, 240)
MINDRAY_DARK = RGBColor(13, 20, 41)
WHITE = RGBColor(255, 255, 255)
LIGHT_GRAY = RGBColor(200, 200, 200)
RED = RGBColor(255, 80, 80)
ORANGE = RGBColor(255, 165, 0)
GREEN = RGBColor(0, 200, 100)

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs

def add_title_slide(prs, title, subtitle=""):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = MINDRAY_DARK
    shape.line.fill.background()
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.333), Inches(1.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = MINDRAY_BLUE
    p.alignment = PP_ALIGN.CENTER
    
    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(12.333), Inches(1))
        tf = sub_box.text_frame
        p = tf.paragraphs[0]
        p.text = subtitle
        p.font.size = Pt(24)
        p.font.color.rgb = LIGHT_GRAY
        p.alignment = PP_ALIGN.CENTER

def add_content_slide(prs, title, bullet_points=None):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = MINDRAY_DARK
    shape.line.fill.background()
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(12), Inches(0.6))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = MINDRAY_BLUE
    
    if bullet_points:
        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(12), Inches(5.8))
        tf = content_box.text_frame
        tf.word_wrap = True
        
        for i, point in enumerate(bullet_points):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            p.text = point
            p.font.size = Pt(18)
            p.font.color.rgb = WHITE
            p.level = 0
            p.space_after = Pt(10)

def add_architecture_slide(prs):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = MINDRAY_DARK
    shape.line.fill.background()
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(12), Inches(0.6))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = '系统整体架构'
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = MINDRAY_BLUE
    
    arch_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(12), Inches(5.5))
    tf = arch_box.text_frame
    tf.word_wrap = True
    
    architecture_text = [
        '[前端可视化大屏] ECharts地图 + 实时数据面板 + 语音播报',
        '          | 每15秒轮询',
        '[数据处理服务层] Python HTTP Server + JSON数据中转',
        '          | 每10分钟爬取',
        '[数据爬取层] Playwright自动化 + MSP系统对接',
        '          |',
        '[源数据系统] MSPS (迈瑞服务支持平台)',
    ]
    
    for i, line in enumerate(architecture_text):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = line
        p.font.size = Pt(20)
        p.font.color.rgb = RGBColor(0, 255, 100) if '[' in line else WHITE
        p.space_after = Pt(8)

def add_table_slide(prs, title, table_data, headers):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = MINDRAY_DARK
    shape.line.fill.background()
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(12), Inches(0.6))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = MINDRAY_BLUE
    
    rows = len(table_data) + 1
    cols = len(headers)
    table_width = Inches(11)
    table_height = Inches(0.45 * rows)
    
    table = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.5), table_width, table_height).table
    
    for j, header in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = MINDRAY_BLUE
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(16)
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER
    
    for i, row_data in enumerate(table_data):
        for j, cell_data in enumerate(row_data):
            cell = table.cell(i + 1, j)
            cell.text = str(cell_data)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(16)
            p.font.color.rgb = WHITE
            if i % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(30, 40, 60)

def add_value_slide(prs, title, metrics):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(1))
    shape.fill.solid()
    shape.fill.fore_color.rgb = MINDRAY_DARK
    shape.line.fill.background()
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(12), Inches(0.6))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = MINDRAY_BLUE
    
    card_width = Inches(3.5)
    card_height = Inches(2.5)
    gap = Inches(0.5)
    
    for i, (metric, value, desc) in enumerate(metrics):
        left = Inches(0.8) + i * (card_width + gap)
        top = Inches(1.8)
        
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, card_width, card_height)
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(20, 30, 50)
        card.line.color.rgb = MINDRAY_BLUE
        
        value_box = slide.shapes.add_textbox(left, top + Inches(0.3), card_width, Inches(1))
        tf = value_box.text_frame
        p = tf.paragraphs[0]
        p.text = value
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = MINDRAY_BLUE
        p.alignment = PP_ALIGN.CENTER
        
        name_box = slide.shapes.add_textbox(left, top + Inches(1.2), card_width, Inches(0.6))
        tf = name_box.text_frame
        p = tf.paragraphs[0]
        p.text = metric
        p.font.size = Pt(18)
        p.font.color.rgb = LIGHT_GRAY
        p.alignment = PP_ALIGN.CENTER
        
        desc_box = slide.shapes.add_textbox(left, top + Inches(1.8), card_width, Inches(0.8))
        tf = desc_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = desc
        p.font.size = Pt(14)
        p.font.color.rgb = GREEN
        p.alignment = PP_ALIGN.CENTER

def main():
    prs = create_presentation()
    
    # 1. 封面
    add_title_slide(prs, '迈瑞中国区技术支持数据中心', '实时监控大屏系统 - 汇报演示')
    
    # 2. 目录
    add_content_slide(prs, '汇报目录', [
        '一、项目背景与目标',
        '二、系统整体架构',
        '三、核心功能展示（6项）',
        '四、技术亮点',
        '五、业务价值与收益',
        '六、后续规划'
    ])
    
    # 3. 项目背景
    add_content_slide(prs, '项目背景与目标', [
        '当前痛点：',
        '  * 数据分散：工单数据散落在MSP系统中',
        '  * 响应滞后：异常工单发现不及时',
        '  * 人力调度难：工程师负载不透明',
        '',
        '建设目标：',
        '  * 实时可视：全国异常工单一屏掌握',
        '  * 智能预警：超时工单自动语音播报',
        '  * 精准定位：乡镇级医院精准地图展示',
        '  * 数据驱动：为管理决策提供实时数据支撑'
    ])
    
    # 4. 系统架构
    add_architecture_slide(prs)
    
    # 5. 核心功能 - 地图
    add_content_slide(prs, '核心功能 1/6：全国工单实时地图', [
        '功能描述：',
        '  * 中国地图底色分区（五大区域着色）',
        '  * 异常工单以动态散点形式标注',
        '  * 点位大小 = 超时严重程度',
        '  * 点位颜色 = 紧急程度',
        '',
        '交互效果：',
        '  * 鼠标悬停可查看详细信息',
        '  * 密集光点直观展示全国热力分布'
    ])
    
    # 6. 核心功能 - 语音
    add_content_slide(prs, '核心功能 2/6：智能语音告警', [
        '功能描述：',
        "  * 点击'语音播报'按钮开启告警",
        '  * 系统检测到更严重超时时自动播报',
        '  * 播报文案：紧急提醒，XX医院工单已超时XX小时...',
        '',
        '设计亮点：',
        '  * 智能触发：基于超时恶化程度触发',
        '  * 防打扰：关闭后彻底安静',
        '  * 中文语音：浏览器原生TTS，无需插件'
    ])
    
    # 7. 核心功能 - 分级
    add_table_slide(prs, '核心功能 3/6：紧急程度三级预警', [
        ['严重超时', '> 72小时', '红色', '呼吸灯闪烁'],
        ['紧急超时', '48-72小时', '橙色', '静态高亮'],
        ['一般超时', '30-48小时', '黄色', '静态显示']
    ], ['级别', '阈值', '颜色', '动画效果'])
    
    # 8. 核心功能 - 负载
    add_content_slide(prs, '核心功能 4/6：工程师负载看板', [
        '功能描述：',
        '  * 自动聚合统计每位工程师异常工单数量',
        '  * 按负载量降序排列，一目了然',
        '',
        '示例展示：',
        '  靳鹏云    3单',
        '  刘志祥    2单',
        '  曹红华    2单',
        '  周海      1单',
        '',
        '管理价值：快速识别谁手里异常单最多，协调支援资源'
    ])
    
    # 9. 核心功能 - 趋势
    add_content_slide(prs, '核心功能 5/6：24小时趋势监控', [
        '功能描述：',
        '  * 折线图展示过去24小时异常工单数量变化',
        '  * 横轴：小时；纵轴：异常工单数',
        '',
        '管理价值：',
        '  * 发现异常高峰时段',
        '  * 评估改进措施效果',
        '  * 为排班和资源配置提供数据支撑'
    ])
    
    # 10. 核心功能 - 滚动条
    add_content_slide(prs, '核心功能 6/6：顶部滚动告警条', [
        '功能描述：',
        '  * 顶部红色横幅滚动显示最紧急的TOP3工单',
        '  * 类似新闻滚动条效果，24小时不间断',
        '',
        '文案示例：',
        '  潜江市人民医院(86h) || 孝感市中心医院(72h)',
        '',
        '价值：即使不看详细面板，也能第一时间获知紧急事件'
    ])
    
    # 11. 技术亮点 - 定位
    add_content_slide(prs, '技术亮点 1/3：精准地理定位', [
        '问题：乡镇级医院地址不含省份关键词，传统匹配无法识别',
        '',
        '解决方案：',
        '  百度地图API地理编码 -> 经纬度坐标 -> 区域范围匹配',
        '',
        '效果对比：',
        '  改造前：深圳市南山区医院 -> 显示在深圳市中心',
        '  改造后：深圳市南山区医院 -> 精确到南山区',
        '',
        '  改造前：潜江市人民医院 -> 无法识别省份',
        '  改造后：潜江市人民医院 -> 正确归属华东地区'
    ])
    
    # 12. 技术亮点 - 多线程
    add_content_slide(prs, '技术亮点 2/3：多线程并发架构', [
        '设计思路：',
        '  * 爬虫线程：独立运行，每10分钟抓取一次',
        '  * 处理线程：监控文件变化，自动生成JSON',
        '  * 前端线程：每15秒轮询JSON，实时更新画面',
        '  * 线程同步：使用threading.Event()确保数据就绪',
        '',
        '优势：',
        '  * 爬虫和前端互不阻塞',
        '  * 数据流水线持续运转',
        '  * 系统稳定性高'
    ])
    
    # 13. 技术亮点 - 区域着色
    add_table_slide(prs, '技术亮点 3/3：五大区域地图着色', [
        ['华东地区', '深蓝', '湖北、上海、江苏、浙江等'],
        ['华北地区', '深紫', '北京、天津、河北、河南等'],
        ['华南地区', '深绿', '广东、广西、海南'],
        ['西南地区', '深橙', '四川、重庆、贵州、云南等'],
        ['西北地区', '深红', '陕西、甘肃、新疆、青海等']
    ], ['大区', '颜色', '包含省份'])
    
    # 14. 业务价值
    add_value_slide(prs, '业务价值与收益', [
        ('异常发现时间', '1000x', '从小时级到秒级'),
        ('工单响应', '提前数小时', '语音主动提醒'),
        ('人力调度', '精准匹配', '数据驱动决策')
    ])
    
    # 15. 后续规划
    add_content_slide(prs, '后续规划', [
        '短期优化（1-2周）：',
        '  * 增加一键导出报表功能',
        '  * 添加历史数据持久化存储（SQLite）',
        '  * 优化地图缩放交互（省份点击下钻）',
        '',
        '中期规划（1-2月）：',
        '  * 增加工程师GPS定位',
        '  * 短信/企业微信告警推送',
        '',
        '长期愿景：',
        '  打造迈瑞全国技术服务智能指挥中心'
    ])
    
    # 16. 结尾
    add_title_slide(prs, '感谢聆听', '迈瑞中国区技术支持数据中心 - 实时监控 - 智能预警 - 数据驱动')
    
    output_file = '迈瑞技术支持数据中心_汇报演示.pptx'
    prs.save(output_file)
    print(f'PPT生成成功：{output_file}')
    print(f'共 {len(prs.slides)} 页幻灯片')

if __name__ == '__main__':
    main()
