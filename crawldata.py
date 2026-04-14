import os
from playwright.sync_api import sync_playwright, Browser, BrowserContext, Page
import time
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import sys
from datetime import datetime, timedelta

# ====================== 配置区域 ======================
WORKORDER_REPORT_URL = "https://msps.mindray.com/#/W3sidHlwZSI6IkwyIiwidmlld05hbWUiOiJyZXBvcnQucmVwb3J0P3JlcG9ydD9yZXBvcnRDb2RlPXF1YWxpdHlJbXByb3ZlIiwidGFiTmFtZSI6Iui0qOmHj+S4iuWNhyIsImFjdGl2ZSI6ZmFsc2UsIm1lbnVDb2RlIjoibTM4MDIiLCJyb3V0ZVBhcmFtIjp7InJlcG9ydENvZGUiOiJxdWFsaXR5SW1wcm92ZSJ9fSx7InR5cGUiOiJMMiIsInZpZXdOYW1lIjoicmVwb3J0LnJlcG9ydD9yZXBvcnRDb2RlPWFjdGl2aXR5aGVhZGVycmVwb3J0IiwidGFiTmFtZSI6IuW3peWNleaKpeihqCIsImFjdGl2ZSI6dHJ1ZSwibWVudUNvZGUiOiJwMjkwMSIsInJvdXRlUGFyYW0iOnsicmVwb3J0Q29kZSI6ImFjdGl2aXR5aGVhZGVycmVwb3J0In19XQ=="

DATA_FOLDER = "工单报表数据"


# ======================================================

def error_exit(message):
    """报错并终止程序"""
    print(f"❌ 错误: {message}")
    sys.exit(1)


def write_log(message):
    """写入日志文件"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_message = f"[{timestamp}] {message}"
    print(log_message)

    # 使用相对路径，兼容Mac和Windows
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(script_dir, "工单报表爬取日志.txt")
    try:
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(log_message + "\n")
    except Exception as e:
        print(f"⚠️ 写入日志失败: {e}")

def create_data_folder():
    """创建数据文件夹"""
    if not os.path.exists(DATA_FOLDER):
        os.makedirs(DATA_FOLDER)
        print(f"✅ 创建数据文件夹: {DATA_FOLDER}")
        write_log(f"创建数据文件夹: {DATA_FOLDER}")
    return DATA_FOLDER


def wait_for_page_stable(page, timeout=30):
    """等待页面稳定（没有网络请求）"""
    try:
        page.wait_for_load_state('networkidle', timeout=timeout * 1000)
        return True
    except Exception as e:
        write_log(f"⚠️ 页面稳定等待超时: {e}")
        return False


def check_login_status(page):
    """检查登录状态，返回True表示已登录，False表示未登录"""
    try:
        current_url = page.url
        title = page.title()

        write_log(f"当前URL: {current_url}")
        write_log(f"页面标题: {title}")

        # 检查是否包含登录相关关键词
        if "login" in current_url.lower() or "登录" in title:
            return False

        # 尝试查找登录表单元素
        login_indicators = [
            "input[type='password']",
            "input[placeholder*='密码']",
            "button:has-text('登录')",
            ".login-form",
            "#loginForm"
        ]

        for selector in login_indicators:
            try:
                if page.locator(selector).first.count() > 0:
                    write_log(f"检测到登录元素: {selector}")
                    return False
            except:
                continue

        return True
    except Exception as e:
        write_log(f"⚠️ 检查登录状态失败: {e}")
        return True  # 出错时假设已登录，让用户手动确认


def remove_blank_first_row(headers, all_data):
    """
    删除空白的第一行数据

    参数:
        headers: 表头列表
        all_data: 所有数据行

    返回:
        (清理后的headers, 清理后的all_data)
    """
    if not all_data:
        return headers, all_data

    # 检查第一行是否全为空或空白
    first_row = all_data[0]
    is_blank = all(cell.strip() == '' for cell in first_row)

    if is_blank:
        write_log("⚠️ 检测到第一行为空白行，正在删除...")
        all_data = all_data[1:]  # 删除第一行
        write_log(f"✅ 已删除空白第一行，剩余 {len(all_data)} 条数据")

    return headers, all_data


def filter_timeout_data(headers, all_data, data_folder):
    """
    筛选超时数据并保存为新表格
    规则：当前系统时间 - 创建时间 > 30小时

    参数:
        headers: 表头列表
        all_data: 所有数据行
        data_folder: 数据文件夹路径
    """
    try:
        write_log("开始筛选超时数据...")

        # 找到"创建时间"列的索引
        create_time_col_index = None
        for i, header in enumerate(headers):
            if "创建时间" in header.strip():
                create_time_col_index = i
                write_log(f"找到'创建时间'列，索引为: {i}")
                break

        if create_time_col_index is None:
            write_log("⚠️ 未找到'创建时间'列，跳过超时数据筛选")
            write_log(f"当前表头: {headers}")
            return

        # 筛选超时数据
        timeout_data = []
        current_time = datetime.now()

        for row_idx, row_data in enumerate(all_data):
            try:
                # 获取创建时间单元格的值
                create_time_str = row_data[create_time_col_index].strip()

                if not create_time_str:
                    continue

                # 尝试解析时间字符串（支持多种常见格式）
                create_time = None
                for fmt in [
                    "%Y-%m-%d %H:%M:%S",
                    "%Y/%m/%d %H:%M:%S",
                    "%Y-%m-%d %H:%M",
                    "%Y/%m/%d %H:%M",
                    "%Y-%m-%d",
                    "%Y/%m/%d"
                ]:
                    try:
                        create_time = datetime.strptime(create_time_str, fmt)
                        break
                    except ValueError:
                        continue

                if create_time is None:
                    write_log(f"⚠️ 第{row_idx + 2}行时间格式无法解析: {create_time_str}")
                    continue

                # 计算时间差（小时）
                time_diff = (current_time - create_time).total_seconds() / 3600

                # 如果超过30小时，加入超时数据
                if time_diff > 30:
                    timeout_data.append(row_data)

            except Exception as e:
                write_log(f"⚠️ 处理第{row_idx + 2}行时出错: {e}")
                continue

        write_log(f"共筛选出 {len(timeout_data)} 条超时数据")

        # 如果有超时数据，保存为新表格
        if timeout_data:
            timeout_filename = "工单报表_超时数据.xlsx"
            timeout_file = os.path.join(data_folder, timeout_filename)

            wb = Workbook()
            ws = wb.active
            ws.title = "超时工单数据"

            # 设置样式
            header_font = Font(bold=True, color="FFFFFF", size=11)
            header_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
            thin_border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )

            # 写入表头
            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_num, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            # 写入数据
            for row_num, row_data in enumerate(timeout_data, 2):
                for col_num, value in enumerate(row_data, 1):
                    cell = ws.cell(row=row_num, column=col_num, value=value)
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    cell.border = thin_border

            # 设置列宽
            for i in range(1, min(len(headers) + 1, 27)):
                ws.column_dimensions[chr(64 + i)].width = 18

            ws.freeze_panes = 'A2'

            # 保存文件
            wb.save(timeout_file)
            write_log(f"✅ 超时数据已保存: {timeout_file}")
            write_log(f"✅ 共导出 {len(timeout_data)} 条超时数据")
        else:
            write_log("ℹ️ 没有超时的工单数据")

    except Exception as e:
        write_log(f"❌ 筛选超时数据失败: {e}")
        import traceback
        write_log(traceback.format_exc())


class WorkOrderCrawler:
    """工单报表爬虫类 - 支持浏览器复用"""

    def __init__(self):
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self.is_initialized = False

    def init_browser(self):
        """初始化浏览器（仅调用一次）"""
        if self.is_initialized:
            print("ℹ️ 浏览器已初始化，跳过")
            return

        print("\n【初始化】启动浏览器...")
        write_log("启动浏览器...")

        self.playwright = sync_playwright().start()

        try:
            self.browser = self.playwright.chromium.launch(headless=False, channel="msedge")
            print("✅ 使用 Microsoft Edge 浏览器")
        except Exception as e:
            print(f"Edge 不可用，使用默认 Chromium: {e}")
            self.browser = self.playwright.chromium.launch(headless=False)

        self.context = self.browser.new_context(viewport={"width": 1920, "height": 1080})
        self.page = self.context.new_page()
        self.is_initialized = True
        write_log("✅ 浏览器已启动")

    def close(self):
        """关闭浏览器"""
        try:
            if self.browser:
                self.browser.close()
                write_log("浏览器已关闭")
            if self.playwright:
                self.playwright.stop()
        except Exception as e:
            write_log(f"⚠️ 关闭浏览器时出错: {e}")

    def crawl(self, first_run=True):
        """
        执行一次数据抓取

        参数:
            first_run: 是否为首次运行（需要手动确认）
        """
        if not self.is_initialized:
            raise Exception("浏览器未初始化，请先调用 init_browser()")

        print("=" * 60)
        print("开始执行工单报表爬取任务...")
        print("=" * 60)

        # 创建数据文件夹
        data_folder = create_data_folder()

        try:
            # ==================== 第一步：访问页面并检查登录状态 ====================
            print("\n【步骤1】访问工单报表页面...")
            write_log("开始访问工单报表页面...")

            try:
                # 如果页面不存在或已关闭，重新导航
                try:
                    self.page.title()
                except:
                    write_log("⚠️ 页面已失效，重新创建页面...")
                    self.page = self.context.new_page()

                self.page.goto(WORKORDER_REPORT_URL, timeout=120000, wait_until="domcontentloaded")
                write_log("✅ 页面初步加载完成")
            except Exception as e:
                write_log(f"⚠️ 页面加载超时或出错: {e}")
                write_log("尝试继续执行...")

            # 等待页面稳定
            write_log("等待页面稳定...")
            wait_for_page_stable(self.page, timeout=30)
            time.sleep(3)

            # 检查是否已登录
            is_logged_in = check_login_status(self.page)

            if not is_logged_in:
                print("⚠️ 检测到未登录状态")
                print("⚠️ 请在浏览器中手动完成登录后，按回车键继续...")
                print("💡 提示：登录状态将被保持，后续抓取无需重新登录")
                write_log("等待用户手动登录...")

                try:
                    input("\n👉 登录完成后，按回车键继续...")
                    print("✅ 用户已确认登录完成")
                    write_log("用户手动登录完成")

                    # 登录后等待页面稳定
                    time.sleep(2)
                    wait_for_page_stable(self.page, timeout=20)
                except Exception as e:
                    write_log(f"⚠️ 等待用户输入时出错: {e}")
                    print("⚠️ 检测到页面变化，尝试继续...")
            else:
                print("✅ 页面已加载，无需登录")
                write_log("页面已加载，无需登录")

            time.sleep(2)

            # ==================== 第二步：等待用户手动设置筛选条件（仅首次运行）====================
            if first_run:
                print("\n" + "=" * 60)
                print("【步骤2】请手动设置筛选条件")
                print("=" * 60)
                print("\n请在浏览器中完成以下操作：")
                print("1. 点击切换到'工单报表'标签页")
                print("2. 设置所需的筛选条件（可选）")
                print("3. 准备就绪后按回车键，程序将自动执行抓取")
                print("=" * 60)
                print("\n💡 提示：前端界面将在新标签页打开，不会覆盖此窗口")
                print("💡 提示：MSP浏览器将保持打开，后续抓取无需重新登录")

                input("\n👉 准备好后，按回车键开始抓取...")
                write_log("用户确认开始抓取（首次运行）")
                print("✅ 开始抓取数据...")
            else:
                print("\n【步骤2】自动模式 - 跳过手动确认")
                write_log("自动模式 - 跳过手动确认（非首次运行）")
                print("✅ 直接开始抓取数据...")

            time.sleep(2)

            # ==================== 第三步：执行单次数据抓取 ====================
            write_log(f"\n{'=' * 60}")
            write_log(f"开始数据抓取")
            write_log(f"{'=' * 60}")

            try:
                # 检查页面是否仍然可用
                try:
                    self.page.title()
                except Exception as e:
                    write_log(f"⚠️ 页面已失效: {e}")
                    write_log("尝试重新加载页面...")
                    try:
                        self.page.reload(timeout=60000, wait_until="domcontentloaded")
                        wait_for_page_stable(self.page, timeout=20)
                        time.sleep(2)
                    except Exception as reload_error:
                        write_log(f"❌ 重新加载页面失败: {reload_error}")
                        raise

                # 点击执行按钮
                write_log("正在自动点击'执行'按钮...")
                try:
                    # 先等待执行按钮可见，最多等待10秒
                    execute_button = self.page.locator(
                        "button:has-text('执行'), button:has-text('查询'), .el-button:has-text('执行')").first

                    # 等待按钮变为可见状态
                    execute_button.wait_for(state="visible", timeout=10000)

                    if execute_button.count() > 0:
                        execute_button.click()
                        write_log("✅ 已点击执行按钮")
                    else:
                        write_log("⚠️ 未找到执行按钮，跳过点击")
                except Exception as e:
                    write_log(f"⚠️ 点击执行按钮失败: {e}")
                    write_log("尝试使用备用定位方式...")
                    try:
                        # 备用方案：尝试点击所有可能的按钮
                        backup_selectors = [
                            "button.el-button--primary",
                            ".query-btn",
                            "[class*='execute']",
                            "[class*='query']"
                        ]
                        for selector in backup_selectors:
                            try:
                                btn = self.page.locator(selector).first
                                if btn.count() > 0 and btn.is_visible():
                                    btn.click()
                                    write_log(f"✅ 使用备用定位器点击成功: {selector}")
                                    break
                            except:
                                continue
                    except Exception as backup_error:
                        write_log(f"❌ 备用方案也失败: {backup_error}")

                # 等待10秒让数据加载
                write_log("等待10秒让数据加载...")
                for i in range(10, 0, -1):
                    time.sleep(1)
                    if i % 2 == 0:
                        write_log(f"  剩余 {i} 秒...")

                # 等待表格加载
                write_log("等待表格加载...")
                time.sleep(2)
                try:
                    self.page.locator("table").first.wait_for(state="visible", timeout=10000)
                    write_log("✅ 表格已加载")
                except:
                    write_log("⚠️ 表格加载超时")

                # 爬取数据
                write_log("开始爬取数据...")
                all_data = []
                headers = []
                skip_first_data_row = False

                # 获取表头
                try:
                    header_elements = self.page.locator("table thead th").all()
                    if not header_elements:
                        header_elements = self.page.locator("th").all()

                    headers = [th.text_content().strip() for th in header_elements if th.text_content().strip()]

                    # 如果还是没有获取到表头，尝试从tbody的第一行获取
                    if not headers:
                        write_log("⚠️ 未能从thead获取表头，尝试从tbody第一行获取...")
                        first_row = self.page.locator("table tbody tr").first
                        if first_row.count() > 0:
                            header_cells = first_row.locator("td").all()
                            headers = [cell.text_content().strip() for cell in header_cells if
                                       cell.text_content().strip()]
                            write_log(f"✅ 从tbody第一行获取到表头: {headers}")

                            # 如果从tbody第一行获取到了表头，需要将这一行从数据中排除
                            skip_first_data_row = True
                        else:
                            write_log("❌ 无法从任何位置获取表头")
                    else:
                        write_log(f"✅ 从thead获取到表头: {headers}")
                        skip_first_data_row = False

                except Exception as e:
                    write_log(f"⚠️ 获取表头失败: {e}")
                    headers = ["工单号", "持续时长/h", "工单状态", "业务描述", "工单类型",
                               "服务类别", "服务子子类", "受理来源", "申告渠道", "提单人",
                               "VIP客户等级", "工程师级别", "工作方式"]
                    skip_first_data_row = False

                # 获取数据行
                try:
                    rows = self.page.locator("table tbody tr").all()
                    write_log(f"找到 {len(rows)} 行数据")

                    for row_idx, row in enumerate(rows):
                        try:
                            # 如果标记了需要跳过第一行，则跳过
                            if skip_first_data_row and row_idx == 0:
                                write_log("⚠️ 跳过第一行（已作为表头使用）")
                                continue

                            cells = row.locator("td").all()
                            row_data = [cell.text_content().strip() for cell in cells]

                            if not row_data or len(row_data) < 2:
                                continue

                            all_data.append(row_data)

                        except Exception as e:
                            write_log(f"⚠️ 解析行失败: {e}")
                            continue

                    write_log(f"✅ 共抓取 {len(all_data)} 条数据")

                except Exception as e:
                    write_log(f"❌ 爬取数据失败: {e}")

                # 删除空白的第一行数据（如果有的话）
                headers, all_data = remove_blank_first_row(headers, all_data)

                # 导出Excel到数据文件夹（固定文件名，覆盖保存）
                if all_data:
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "工单报表数据"

                    header_font = Font(bold=True, color="FFFFFF", size=11)
                    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                    thin_border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )

                    for col_num, header in enumerate(headers, 1):
                        cell = ws.cell(row=1, column=col_num, value=header)
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.border = thin_border

                    for row_num, row_data in enumerate(all_data, 2):
                        for col_num, value in enumerate(row_data, 1):
                            cell = ws.cell(row=row_num, column=col_num, value=value)
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                            cell.border = thin_border

                    for i in range(1, min(len(headers) + 1, 27)):
                        ws.column_dimensions[chr(64 + i)].width = 18

                    ws.freeze_panes = 'A2'

                    # 固定文件名（不含时间戳和次数）
                    fixed_filename = "工单报表_最新数据.xlsx"
                    output_file = os.path.join(data_folder, fixed_filename)

                    # 使用固定文件名保存（覆盖之前的文件）
                    wb.save(output_file)
                    write_log(f"✅ Excel已保存: {output_file}")
                    write_log(f"✅ 本次共导出 {len(all_data)} 条数据")

                    # 筛选并保存超时数据
                    filter_timeout_data(headers, all_data, data_folder)
                else:
                    write_log("⚠️ 没有找到数据")

                write_log(f"数据抓取完成")
                print(f"\n✅ 第 {'首' if first_run else ''}次数据抓取完成")
                print(f"💡 MSP浏览器保持打开状态，下次抓取将复用")



            except Exception as e:
                write_log(f"❌ 数据抓取发生错误: {e}")
                import traceback
                write_log(traceback.format_exc())
                raise

        except KeyboardInterrupt:
            write_log("\n⚠️ 用户中断程序")
            print("\n⚠️ 程序被用户中断")
            raise
        except Exception as e:
            write_log(f"❌ 爬取任务发生错误: {e}")
            import traceback
            write_log(traceback.format_exc())
            raise

def crawl_workorder_report(first_run=True, crawler=None):
    """
    兼容旧接口的函数

    参数:
        first_run: 是否为首次运行
        crawler: WorkOrderCrawler实例（可选）
    """
    if crawler is None:
        # 如果没有提供crawler实例，创建临时实例（旧行为）
        temp_crawler = WorkOrderCrawler()
        temp_crawler.init_browser()
        try:
            temp_crawler.crawl(first_run=first_run)
        finally:
            temp_crawler.close()
    else:
        # 使用提供的crawler实例
        crawler.crawl(first_run=first_run)


if __name__ == "__main__":
    crawl_workorder_report()
