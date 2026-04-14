import os
import sys
import time
import threading
from datetime import datetime

def get_resource_path(relative_path):
    """获取资源文件的绝对路径（支持打包后的exe）"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# 导入其他模块的功能
from crawldata import crawl_workorder_report, DATA_FOLDER, WorkOrderCrawler
from getHospital import extract_workorder_data, generate_hospital_json, find_available_port, start_http_server, \
    open_browser


def run_data_crawler(data_ready_event, data_updated_event):
    """
    运行数据爬取任务（在后台线程中持续运行）
    """
    print("\n" + "=" * 60)
    print("【数据爬取模块】已启动")
    print("=" * 60)

    # 创建持久的爬虫实例
    crawler = WorkOrderCrawler()

    try:
        # 初始化浏览器（只执行一次）
        crawler.init_browser()
        print("✅ 浏览器已初始化，将在所有抓取任务中复用")

        crawl_count = 0

        while True:
            crawl_count += 1
            print(f"\n{'=' * 60}")
            print(f"开始第 {crawl_count} 次数据爬取...")
            print(f"{'=' * 60}")

            # 运行一次完整的爬取流程（仅首次需要手动确认）
            is_first_run = (crawl_count == 1)
            crawler.crawl(first_run=is_first_run)

            # 如果是首次爬取，通知主线程：数据已准备就绪
            if crawl_count == 1:
                data_ready_event.set()
                print("\n✓ 首次数据爬取完成，通知前端服务启动")

            # 通知数据处理模块有新数据
            data_updated_event.set()
            print(f"\n✓ 第 {crawl_count} 次数据爬取完成，数据已更新")

            # 等待10分钟后进行下一次抓取
            print(f"\n等待10分钟后进行第 {crawl_count + 1} 次抓取...")
            for i in range(600, 0, -1):
                time.sleep(1)
                if i % 60 == 0:
                    print(f"  下次抓取倒计时: {i // 60} 分钟...")

    except KeyboardInterrupt:
        print("\n数据爬取模块已停止")
        crawler.close()
        if crawl_count == 0:
            data_ready_event.set()  # 即使出错也通知前端启动
    except Exception as e:
        print(f"\n❌ 数据爬取模块异常: {e}")
        import traceback
        traceback.print_exc()
        crawler.close()
        if crawl_count == 0:
            data_ready_event.set()  # 即使出错也通知前端启动



def run_data_processor(html_file, data_ready_event, data_updated_event):
    """
    处理数据并启动HTTP服务器（在后台线程中运行）
    """
    # 等待数据爬取完成
    print("\n等待首次数据爬取完成...")
    data_ready_event.wait()
    print("✓ 数据已就绪，开始启动前端服务")

    print("\n" + "=" * 60)
    print("【数据处理和前端服务模块】已启动")
    print("=" * 60)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_file = os.path.join(script_dir, DATA_FOLDER, "工单报表_最新数据.xlsx")
    json_output = os.path.join(script_dir, "hospital_data.json")

    # 检查HTML文件是否存在
    if not os.path.exists(html_file):
        print(f"❌ 错误：找不到HTML文件 {html_file}")
        return

    # 先处理一次数据
    print("\n步骤 1/3: 正在处理首次爬取的数据...")
    try:
        if os.path.exists(excel_file):
            workorder_list = extract_workorder_data(excel_file)

            if workorder_list:
                print(f"✓ 成功提取 {len(workorder_list)} 条工单数据")

                # 生成JSON数据文件
                map_data = generate_hospital_json(workorder_list, json_output)
                print(f"✓ JSON数据已生成: {json_output}")

                # 打印生成的数据统计
                type_stats = {}
                for item in map_data:
                    type_stats[item['type']] = type_stats.get(item['type'], 0) + 1

                print("\n数据类型统计：")
                type_names = {
                    'normal': '正常工单',
                    'timeout': '超时服务',
                    'complaint': '投诉工单',
                    'project': '大项目交付'
                }
                for type_key, count in type_stats.items():
                    print(f"  {type_names.get(type_key, type_key)}: {count} 条")
            else:
                print("⚠️ 未提取到任何数据")
        else:
            print(f"⚠️ 数据文件不存在: {excel_file}")
    except Exception as e:
        print(f"❌ 首次数据处理失败: {e}")
        import traceback
        traceback.print_exc()

    # 启动HTTP服务器
    print("\n步骤 2/3: 正在启动HTTP服务器...")
    try:
        port = find_available_port(8080)
        print(f"✓ 使用端口: {port}")

        # 在新线程中启动HTTP服务器
        server_thread = threading.Thread(target=start_http_server, args=(port, script_dir), daemon=True)
        server_thread.start()

        # 等待服务器启动
        time.sleep(2)

        # 打开浏览器（直接在当前线程调用，不使用额外线程）
        print("\n步骤 3/3: 正在打开浏览器...")
        open_browser(port, delay=0)  # 立即打开，不再延迟

        print("\n" + "=" * 60)
        print("✓ 前端服务已启动！")
        print(f"✓ 访问地址: http://localhost:{port}/dashboard.html")
        print("=" * 60)

        # 持续处理数据并更新JSON
        print("\n【数据自动更新】已启动（等待新数据到达）")
        while True:
            try:
                # 等待数据更新信号
                data_updated_event.wait()
                data_updated_event.clear()  # 重置事件

                print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 检测到新数据，正在处理...")

                # 检查Excel文件是否存在
                if os.path.exists(excel_file):
                    # 提取数据
                    workorder_list = extract_workorder_data(excel_file)

                    if workorder_list:
                        print(f"✓ 成功提取 {len(workorder_list)} 条工单数据")

                        # 生成JSON数据文件
                        map_data = generate_hospital_json(workorder_list, json_output)
                        print(f"✓ JSON数据已更新: {json_output}")
                    else:
                        print("⚠️ 未提取到任何数据")
                else:
                    print(f"⚠️ 数据文件不存在: {excel_file}")

            except Exception as e:
                print(f"❌ 数据处理异常: {e}")
                import traceback
                traceback.print_exc()
                time.sleep(5)

    except Exception as e:
        print(f"\n❌ 前端服务启动失败: {e}")
        import traceback
        traceback.print_exc()


def main():
    """
    主程序：统一管理数据爬取和前端展示
    """
    print("\n" + "=" * 60)
    print("迈瑞技术支持驾驶舱 - 自动化数据系统")
    print("=" * 60)
    print(f"\n启动时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("\n系统功能：")
    print("1. 自动登录并爬取工单报表数据")
    print("2. 实时处理数据并生成JSON")
    print("3. 启动HTTP服务器并展示前端")
    print("4. 持续运行，自动刷新数据")
    print("\n提示：按 Ctrl+C 可停止整个系统\n")

    # 获取HTML文件路径（支持打包后的exe）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    html_file = get_resource_path("dashboard.html")

    # 创建数据文件夹
    if not os.path.exists(DATA_FOLDER):
        os.makedirs(DATA_FOLDER)
        print(f"✅ 创建数据文件夹: {DATA_FOLDER}")

    # 创建同步事件：用于等待首次数据爬取完成
    data_ready_event = threading.Event()
    # 创建数据更新事件：用于通知前端有新数据
    data_updated_event = threading.Event()

    # 启动数据爬取线程
    crawler_thread = threading.Thread(target=run_data_crawler, args=(data_ready_event, data_updated_event), daemon=True)
    crawler_thread.start()
    print("✓ 数据爬取线程已启动")

    # 启动数据处理和前端服务线程（会等待数据就绪）
    processor_thread = threading.Thread(target=run_data_processor,
                                        args=(html_file, data_ready_event, data_updated_event), daemon=True)
    processor_thread.start()
    print("✓ 数据处理和前端服务线程已启动（等待首次数据爬取完成）")

    print("\n" + "=" * 60)
    print("✓ 系统已全部启动！")
    print("=" * 60)
    print("\n运行状态：")
    print("- 数据爬取：后台持续运行（每10分钟抓取一次）")
    print("- 数据处理：首次数据完成后启动前端，之后有新数据时立即处理")
    print("- 前端展示：数据就绪后自动打开浏览器，自动刷新")
    print("\n请按 Ctrl+C 停止系统...\n")

    # 主线程保持运行
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\n" + "=" * 60)
        print("系统正在停止...")
        print("=" * 60)
        print("✓ 数据爬取模块已停止")
        print("✓ 数据处理模块已停止")
        print("✓ HTTP服务器已停止")
        print("\n感谢使用迈瑞中国区技术支持数据中心系统！")


if __name__ == "__main__":
    main()



