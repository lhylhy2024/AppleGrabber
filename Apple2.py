import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import os
import sys
import re
from datetime import datetime


def setup_chrome_driver():
    """设置 ChromeDriver - 增强版本"""
    chrome_options = Options()

    # 添加常用参数
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--start-maximized")

    # 用户代理设置
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

    print("正在初始化 ChromeDriver...")

    # 策略1: 尝试使用系统PATH中的chromedriver
    try:
        print("尝试方法1: 使用系统PATH中的ChromeDriver...")
        driver = webdriver.Chrome(options=chrome_options)
        print("成功使用系统PATH中的ChromeDriver")
        return driver
    except Exception as e:
        print(f"方法1失败: {str(e)[:100]}")

    # 策略2: 尝试使用webdriver_manager自动管理
    try:
        print("尝试方法2: 使用webdriver_manager自动下载ChromeDriver...")
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("成功使用webdriver_manager安装ChromeDriver")
        return driver
    except Exception as e:
        print(f"方法2失败: {str(e)[:100]}")

    # 策略3: 尝试查找本地可能存在的chromedriver
    print("尝试方法3: 查找本地可能存在的ChromeDriver...")

    # 定义可能的chromedriver路径
    possible_paths = []

    if sys.platform.startswith('win'):
        # Windows路径
        possible_paths = [
            os.path.join(os.environ.get('USERPROFILE', ''), 'chromedriver.exe'),
            os.path.join(os.getcwd(), 'chromedriver.exe'),
            'chromedriver.exe',
            r'C:\chromedriver.exe',
        ]
    elif sys.platform.startswith('darwin'):
        # macOS路径
        possible_paths = [
            '/usr/local/bin/chromedriver',
            '/Applications/chromedriver',
            os.path.join(os.getcwd(), 'chromedriver'),
            'chromedriver',
        ]
    else:
        # Linux路径
        possible_paths = [
            '/usr/bin/chromedriver',
            '/usr/local/bin/chromedriver',
            os.path.join(os.getcwd(), 'chromedriver'),
            'chromedriver',
        ]

    for path in possible_paths:
        if os.path.exists(path):
            print(f"找到ChromeDriver: {path}")
            try:
                service = Service(executable_path=path)
                driver = webdriver.Chrome(service=service, options=chrome_options)
                print(f"成功使用找到的ChromeDriver: {path}")
                return driver
            except Exception as e:
                print(f"使用{path}失败: {str(e)[:100]}")

    # 策略4: 使用默认安装位置的Chrome
    print("尝试方法4: 使用默认位置的Chrome...")
    try:
        # 尝试使用默认位置
        if sys.platform.startswith('win'):
            chrome_paths = [
                r'C:\Program Files\Google\Chrome\Application\chrome.exe',
                r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
            ]
            for chrome_path in chrome_paths:
                if os.path.exists(chrome_path):
                    chrome_options.binary_location = chrome_path
                    break
        elif sys.platform.startswith('darwin'):
            chrome_path = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
            if os.path.exists(chrome_path):
                chrome_options.binary_location = chrome_path

        driver = webdriver.Chrome(options=chrome_options)
        print("成功使用默认位置的Chrome")
        return driver
    except Exception as e:
        print(f"方法4失败: {str(e)[:100]}")

    print("\n" + "=" * 60)
    print("ChromeDriver启动失败，请按以下步骤操作：")
    print("1. 确保已安装Chrome浏览器")
    print("2. 下载对应版本的ChromeDriver: https://chromedriver.chromium.org/")
    print("3. 将chromedriver放在以下位置之一：")
    for path in possible_paths:
        print(f"   - {path}")
    print("4. 或添加到系统PATH环境变量中")
    print("=" * 60)

    return None


def extract_developer_id(developer_element):
    """尝试从开发商元素中提取开发商ID"""
    try:
        # 方法1: 查找可能的ID属性
        id_attrs = developer_element.get_attribute("outerHTML")

        # 查找常见的ID模式
        patterns = [
            r'data-developer-id="([^"]+)"',
            r'data-publisher-id="([^"]+)"',
            r'data-seller-id="([^"]+)"',
            r'id="([A-Z0-9]{10,})"',
            r'class="([A-Z0-9]{10,})"',
        ]

        for pattern in patterns:
            match = re.search(pattern, id_attrs, re.IGNORECASE)
            if match:
                return match.group(1)

        # 方法2: 检查文本是否为ID格式
        text = developer_element.text.strip()
        if re.match(r'^[A-Z0-9]{8,}$', text) and not re.search(r'[a-z]', text):
            return text

        # 方法3: 查找子元素中的ID
        child_elements = developer_element.find_elements(By.XPATH, ".//*")
        for child in child_elements:
            child_text = child.text.strip()
            if re.match(r'^[A-Z0-9]{8,}$', child_text) and not re.search(r'[a-z]', child_text):
                return child_text

        return "N/A"
    except Exception as e:
        return "N/A"


def clean_app_name(app_name):
    """清理应用名称"""
    if not app_name or app_name == "N/A":
        return "N/A"

    # 移除多余的空白字符
    app_name = ' '.join(app_name.split())

    # 如果包含"aria-label"等前缀，尝试提取真正的应用名
    if 'aria-label:' in app_name.lower():
        # 尝试提取冒号后的内容
        parts = app_name.split(':', 1)
        if len(parts) > 1:
            return parts[1].strip()

    return app_name


def scrape_apple_report_page():
    """
    根据提供的HTML结构爬取苹果报告问题页面数据
    """

    driver = setup_chrome_driver()
    if not driver:
        print("无法启动 ChromeDriver，请检查上述说明")
        return None

    try:
        print("正在打开页面...")
        driver.get("https://reportaproblem.apple.com/")

        print("\n" + "=" * 60)
        print("重要提示：")
        print("1. 请手动登录您的 Apple ID")
        print("2. 登录成功后，页面会显示购买记录")
        print("3. 确保页面完全加载后，按回车键继续")
        print("=" * 60 + "\n")

        input("登录完成后按回车键继续...")

        # 等待页面加载
        print("正在等待页面加载...")
        time.sleep(5)

        # 保存页面源代码以便调试
        page_source = driver.page_source
        with open("apple_page_detailed.html", "w", encoding="utf-8") as f:
            f.write(page_source)
        print("已保存页面源代码到 apple_page_detailed.html")

        print("正在解析购买记录...")

        # 根据提示词，查找所有购买记录容器
        purchase_containers = driver.find_elements(By.CSS_SELECTOR, "div.purchase.loaded.collapsed")

        if not purchase_containers:
            # 尝试其他可能的类名
            purchase_containers = driver.find_elements(By.CSS_SELECTOR, "div.purchase")
            print(f"找到 {len(purchase_containers)} 个 div.purchase 元素")
        else:
            print(f"找到 {len(purchase_containers)} 个 div.purchase.loaded.collapsed 元素")

        data = []
        record_index = 1

        for purchase_idx, purchase in enumerate(purchase_containers, 1):
            try:
                # 提取购买时间
                purchase_date = "N/A"
                try:
                    date_elements = purchase.find_elements(By.CSS_SELECTOR, "span.invoice-date")
                    if date_elements:
                        purchase_date = date_elements[0].text.strip()
                        print(f"购买记录 {purchase_idx} 购买时间: {purchase_date}")
                except Exception as e:
                    print(f"提取购买时间失败: {e}")

                # 查找该购买记录下的所有应用项目
                app_items = purchase.find_elements(By.CSS_SELECTOR, "li.pli")

                if not app_items:
                    # 尝试其他选择器
                    app_items = purchase.find_elements(By.CSS_SELECTOR, ".pli")
                    print(f"购买记录 {purchase_idx} 找到 {len(app_items)} 个 li.pli 元素")
                else:
                    print(f"购买记录 {purchase_idx} 找到 {len(app_items)} 个应用项目")

                for app_idx, app_item in enumerate(app_items, 1):
                    try:
                        # 提取应用名称
                        app_name = "N/A"
                        try:
                            # 首先尝试从 aria-label 属性获取
                            title_elements = app_item.find_elements(By.CSS_SELECTOR, "div.pli-title")
                            if title_elements:
                                title_element = title_elements[0]
                                # 尝试获取 aria-label 属性
                                aria_label = title_element.get_attribute("aria-label")
                                if aria_label and aria_label.strip():
                                    app_name = aria_label.strip()
                                else:
                                    # 如果没有 aria-label，获取文本内容
                                    app_name = title_element.text.strip()

                            # 如果仍然没有获取到，尝试从图片的 alt 属性获取
                            if app_name == "N/A" or not app_name:
                                img_elements = app_item.find_elements(By.CSS_SELECTOR, ".pli-artwork img")
                                if img_elements:
                                    alt_text = img_elements[0].get_attribute("alt")
                                    if alt_text and alt_text.strip():
                                        app_name = alt_text.strip()
                        except Exception as e:
                            print(f"提取应用名称失败: {e}")

                        app_name = clean_app_name(app_name)

                        # 提取开发商
                        developer = "N/A"
                        developer_id = "N/A"
                        try:
                            developer_elements = app_item.find_elements(By.CSS_SELECTOR,
                                                                        "div.pli-publisher.has-publisher")
                            if not developer_elements:
                                # 尝试没有 has-publisher 类的情况
                                developer_elements = app_item.find_elements(By.CSS_SELECTOR, "div.pli-publisher")

                            if developer_elements:
                                developer_element = developer_elements[0]
                                developer = developer_element.text.strip()

                                # 尝试提取开发商ID
                                developer_id = extract_developer_id(developer_element)

                                # 如果开发商文本本身就是ID格式，交换一下
                                if re.match(r'^[A-Z0-9]{8,}$', developer) and not re.search(r'[a-z]', developer):
                                    if developer_id == "N/A":
                                        developer_id = developer
                                    developer = "N/A"
                        except Exception as e:
                            print(f"提取开发商失败: {e}")

                        # 提取价格
                        price = "N/A"
                        try:
                            price_elements = app_item.find_elements(By.CSS_SELECTOR, "div.pli-price span")
                            if not price_elements:
                                # 尝试直接查找 div.pli-price 的文本
                                price_elements = app_item.find_elements(By.CSS_SELECTOR, "div.pli-price")

                            if price_elements:
                                price = price_elements[0].text.strip()
                                # 如果价格为空，可能是免费应用
                                if not price:
                                    price = "免费"
                        except Exception as e:
                            print(f"提取价格失败: {e}")

                        # 添加到数据列表
                        data.append({
                            "序号": record_index,
                            "软件名称": app_name,
                            "开发商": developer,
                            "开发商ID": developer_id,
                            "购买时间": purchase_date,
                            "价格": price
                        })

                        print(f"  应用 {app_idx}: {app_name} - {developer} (ID: {developer_id}) - {price}")

                        record_index += 1

                    except Exception as e:
                        print(f"处理应用项目失败: {e}")
                        continue

            except Exception as e:
                print(f"处理购买记录失败: {e}")
                continue

        # 创建DataFrame
        if data:
            df = pd.DataFrame(data)

            # 保存到Excel
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"apple_purchases_{timestamp}.xlsx"
            df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"\n数据已保存到: {output_file}")
            print(f"共采集到 {len(data)} 条记录")

            # 显示前几行数据
            print("\n数据预览：")
            print(df.head(10))

            # 统计信息
            print(f"\n统计信息：")
            print(f"  有开发商名称的记录: {len(df[df['开发商'] != 'N/A'])}")
            print(f"  有开发商ID的记录: {len(df[df['开发商ID'] != 'N/A'])}")
            print(f"  免费应用数量: {len(df[df['价格'] == '免费'])}")

            return df
        else:
            print("未能提取到任何购买记录")
            print("\n可能的原因：")
            print("1. 页面结构可能与提示词不同")
            print("2. 没有找到购买记录")
            print("3. 需要等待更长时间让页面加载")

            # 显示页面结构以供调试
            print("\n页面结构分析：")
            analyze_page_structure(driver)

            return None

    except Exception as e:
        print(f"爬取过程中出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

    finally:
        if driver:
            print("\n程序执行完成，5秒后关闭浏览器...")
            time.sleep(5)
            driver.quit()


def analyze_page_structure(driver):
    """分析页面结构以帮助调试"""
    try:
        # 查找所有可能的购买记录元素
        selectors_to_check = [
            "div.purchase",
            "div.transaction",
            "div.order",
            "li.purchase",
            "table",
            "div[class*='purchase']",
            "div[class*='transaction']",
        ]

        print("页面结构分析结果：")
        for selector in selectors_to_check:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            if elements:
                print(f"  找到 {len(elements)} 个 '{selector}' 元素")

                # 显示前3个元素的简要信息
                for i, elem in enumerate(elements[:3]):
                    try:
                        elem_class = elem.get_attribute("class")
                        elem_id = elem.get_attribute("id")
                        text_preview = elem.text[:100] + "..." if len(elem.text) > 100 else elem.text
                        print(f"    元素 {i + 1}: class='{elem_class}', id='{elem_id}', text='{text_preview}'")
                    except:
                        print(f"    元素 {i + 1}: 无法获取信息")

        # 检查是否有应用名称相关的元素
        app_name_selectors = [
            "div.pli-title",
            ".app-name",
            ".software-name",
            ".item-name",
            "[aria-label]"
        ]

        print("\n应用名称相关元素：")
        for selector in app_name_selectors:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            if elements:
                print(f"  找到 {len(elements)} 个 '{selector}' 元素")

        # 检查是否有价格相关的元素
        price_selectors = [
            "div.pli-price",
            ".price",
            ".amount",
            "[class*='price']"
        ]

        print("\n价格相关元素：")
        for selector in price_selectors:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            if elements:
                print(f"  找到 {len(elements)} 个 '{selector}' 元素")
                for i, elem in enumerate(elements[:2]):
                    print(f"    价格 {i + 1}: '{elem.text}'")

    except Exception as e:
        print(f"分析页面结构时出错: {e}")


# 简化的主函数，只尝试自动爬取
def main():
    print("苹果购买记录采集工具")
    print("=" * 60)
    print("根据提示词优化，支持提取开发商ID")
    print("=" * 60)

    # 尝试自动爬取
    result = scrape_apple_report_page()

    if result is not None:
        print("\n程序执行成功！")
        print("=" * 60)
        print("说明：")
        print("1. 如果开发商显示为N/A但开发商ID有值，表示该记录只有ID没有公司名称")
        print("2. 免费应用的价格显示为'免费'")
        print("3. 购买时间从span.invoice-date提取")
    else:
        print("\n自动爬取失败，请检查页面结构和ChromeDriver配置")


if __name__ == "__main__":
    main()