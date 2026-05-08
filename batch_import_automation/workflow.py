"""
批量导入主流程模块"""

import time
import os
import glob
import json
import ctypes
import uiautomation as auto
import logging
import traceback
from pywinauto import Desktop
from pywinauto.findwindows import ElementNotFoundError
from .config import BatchImportConfigManager
from .ui_utils import human_delay, safe_click_control, set_control_value, mouse_click
from .window_handler import WindowHandler
from .field_handler import FieldHandler
from .shop_handler import ShopHandler
from .file_handler import FileHandler
from .template_handler import TemplateHandler
# 导入类目修改中的下拉框处理函数
import sys
import os
# 添加父目录到路径，以便导入category_automation
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)
from category_automation.dropdown_handler import select_dropdown_option_with_scroll
from category_automation.workflow import CategoryAutomation

# Windows API 用于鼠标事件
user32 = ctypes.windll.user32

# 配置日志记录
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('batch_import.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('BatchImportAutomation')


class BatchImportAutomation:
    """批量导入自动化主类"""
    
    WINDOW_NAME = "批量导入"
    
    # 字段名到AutomationId的映射
    FIELD_AUTOMATION_IDS = {
        "关键词": "HotSearchskinTextBox",
        "热搜词": "HotSearchskinTextBox",  # 支持两种名称
        "基础加价": "JiChuskinTextBox",
        "运费": None,  # 需要通过标签查找
        "重量": "MinWeightskinTextBox",  # 重量字段
        "链接": "skinTextBox1",  # 链接输入框
    }
    
    def __init__(self, config_file: str = "属性配置.json",
                 excel_file: str = "批量导入.xlsx",
                 excel_files=None,
                 auto_discover_batch_excels: bool = False):
        """
        初始化自动化类
        :param config_file: 配置文件路径
        :param excel_file: 单个表格路径（excel_files 未指定时使用）
        :param excel_files: 多个表格路径列表，按顺序依次处理；指定后忽略 excel_file 与自动发现
        :param auto_discover_batch_excels: 未指定 excel_files 时，是否在表格所在目录自动收集全部「*批量导入*.xlsx」
        """
        try:
            logger.info(f"初始化BatchImportAutomation，配置文件: {config_file}")
            ad = auto_discover_batch_excels and excel_files is None
            self.config_manager = BatchImportConfigManager(
                config_file,
                excel_file=excel_file,
                excel_files=excel_files,
                auto_discover_batch_excels=ad,
            )
            self.window = None

            # 初始化各个处理器
            self.window_handler = WindowHandler(self.window)
            self.field_handler = FieldHandler(self.window)
            self.shop_handler = ShopHandler(self.window)
            self.file_handler = FileHandler()
            self.template_handler = TemplateHandler(self.window)
            # 初始化类目修改自动化（传入config_manager以便共享Excel配置）
            self.category_automation = CategoryAutomation(config_file=config_file, excel_config_manager=self.config_manager, pause_event=None)
            logger.info("初始化完成")
        except Exception as e:
            logger.error(f"初始化失败: {str(e)}")
            traceback.print_exc()
            raise
    
    def check_pause(self):
        """检查暂停状态（为兼容性保留，当前不实现暂停功能）"""
        # ESC暂停功能已移除，此方法保留以避免代码报错
        pass
    
    def find_window(self, timeout: int = 5) -> bool:
        """
        查找并激活目标窗口
        :param timeout: 超时时间（秒）
        :return: 是否成功找到窗口
        """
        try:
            logger.info(f"查找窗口: {self.WINDOW_NAME}")
            if self.window_handler.find_window(timeout):
                self.window = self.window_handler.window
                # 更新所有处理器的window引用
                self.field_handler.window = self.window
                self.shop_handler.window = self.window
                self.template_handler.window = self.window
                logger.info("窗口查找成功")
                return True
            logger.warning("窗口查找失败")
            return False
        except Exception as e:
            logger.error(f"查找窗口时出错: {str(e)}")
            traceback.print_exc()
            return False
        
    def find_control_by_name(self, name: str, control_type=None):
        """通过名称查找控件（委托给FieldHandler）"""
        try:
            return self.field_handler.find_control_by_name(name, control_type)
        except Exception as e:
            logger.error(f"通过名称查找控件时出错: {str(e)}")
            return None
    
    def find_input_by_label(self, label_text: str):
        """通过标签文本查找对应的输入控件（委托给FieldHandler）"""
        try:
            return self.field_handler.find_input_by_label(label_text)
        except Exception as e:
            logger.error(f"通过标签查找输入控件时出错: {str(e)}")
            return None
    
    def set_field_value(self, field_name: str, value: str) -> bool:
        """设置字段值（委托给FieldHandler）"""
        try:
            return self.field_handler.set_field_value(field_name, value)
        except Exception as e:
            logger.error(f"设置字段值时出错: {str(e)}")
            return False
    
    def select_site(self, site_name: str) -> bool:
        """
        选择站点（使用组合框）
        :param site_name: 站点名称
        :return: 是否成功选择
        """
        try:
            if not self.window:
                logger.warning("窗口未初始化，无法选择站点")
                print("❌ 窗口未初始化，无法选择站点")
                return False
            
            print(f"\n🌐 开始选择站点: {site_name}")
            logger.info(f"选择站点: {site_name}")
            
            # 查找站点组合框（WPF ComboBox）
            site_combo = None
            
            # 方法1: 通过ControlType查找所有ComboBox，然后筛选
            try:
                # 查找所有ComboBox控件
                all_combos = []
                def find_combos(ctrl, depth=0):
                    if depth > 5:
                        return
                    try:
                        if ctrl.ControlType == auto.ControlType.ComboBoxControl:
                            # 检查是否是WPF ComboBox
                            if ctrl.FrameworkId == "WPF" and ctrl.ClassName == "ComboBox":
                                all_combos.append(ctrl)
                    except:
                        pass
                    try:
                        child = ctrl.GetFirstChildControl()
                        while child:
                            find_combos(child, depth + 1)
                            child = child.GetNextSiblingControl()
                    except:
                        pass
                
                find_combos(self.window)
                
                # 通常站点组合框是第一个或第二个ComboBox
                if all_combos:
                    site_combo = all_combos[0]
                    print(f"  ✓ 找到站点组合框（通过遍历查找）")
                    logger.info("找到站点组合框")
            except Exception as e:
                print(f"  ⚠️ 查找组合框时出错: {e}")
                logger.warning(f"查找组合框时出错: {str(e)}")
            
            if not site_combo:
                print("  ⚠️ 未找到站点组合框")
                logger.warning("未找到站点组合框")
                return False
            
            # 使用类目修改中的下拉框处理方式选择站点
            if select_dropdown_option_with_scroll(site_combo, site_name, field_name="站点"):
                print(f"  ✅ 站点选择完成: {site_name}")
                logger.info(f"站点选择完成: {site_name}")
                human_delay(0.5, 0.8)
                return True
            else:
                print(f"  ❌ 站点选择失败: {site_name}")
                logger.warning(f"站点选择失败: {site_name}")
                return False
                
        except Exception as e:
            logger.error(f"选择站点时出错: {str(e)}")
            print(f"  ❌ 选择站点时出错: {e}")
            traceback.print_exc()
            return False
    
    def select_group(self, group_name: str) -> bool:
        """
        选择分组（使用组合框，第二个ComboBox）
        :param group_name: 分组名称
        :return: 是否成功选择
        """
        try:
            if not self.window:
                logger.warning("窗口未初始化，无法选择分组")
                print("❌ 窗口未初始化，无法选择分组")
                return False
            
            print(f"\n📂 开始选择分组: {group_name}")
            logger.info(f"选择分组: {group_name}")
            
            group_combo = None
            
            try:
                all_combos = []
                def find_combos(ctrl, depth=0):
                    if depth > 5:
                        return
                    try:
                        if ctrl.ControlType == auto.ControlType.ComboBoxControl:
                            if ctrl.FrameworkId == "WPF" and ctrl.ClassName == "ComboBox":
                                all_combos.append(ctrl)
                    except:
                        pass
                    try:
                        child = ctrl.GetFirstChildControl()
                        while child:
                            find_combos(child, depth + 1)
                            child = child.GetNextSiblingControl()
                    except:
                        pass
                
                find_combos(self.window)
                
                # 分组组合框是第二个ComboBox（索引1）
                if len(all_combos) >= 2:
                    group_combo = all_combos[1]
                    print(f"  ✓ 找到分组组合框（第2个ComboBox，共找到{len(all_combos)}个）")
                    logger.info("找到分组组合框")
                else:
                    print(f"  ⚠️ ComboBox数量不足（找到{len(all_combos)}个，需要至少2个）")
                    logger.warning(f"ComboBox数量不足: {len(all_combos)}")
            except Exception as e:
                print(f"  ⚠️ 查找组合框时出错: {e}")
                logger.warning(f"查找组合框时出错: {str(e)}")
            
            if not group_combo:
                print("  ⚠️ 未找到分组组合框")
                logger.warning("未找到分组组合框")
                return False
            
            if select_dropdown_option_with_scroll(group_combo, group_name, field_name="分组"):
                print(f"  ✅ 分组选择完成: {group_name}")
                logger.info(f"分组选择完成: {group_name}")
                human_delay(0.5, 0.8)
                return True
            else:
                print(f"  ❌ 分组选择失败: {group_name}")
                logger.warning(f"分组选择失败: {group_name}")
                return False
                
        except Exception as e:
            logger.error(f"选择分组时出错: {str(e)}")
            print(f"  ❌ 选择分组时出错: {e}")
            traceback.print_exc()
            return False
    
    def select_shops(self, shop_names: str, old_shop: str = "", is_site_changed: bool = False, select_all_only: bool = False) -> bool:
        """在列表视图中选择店铺（委托给ShopHandler）"""
        try:
            return self.shop_handler.select_shops(shop_names, old_shop, is_site_changed, select_all_only)
        except Exception as e:
            logger.error(f"选择店铺时出错: {str(e)}")
            return False
    
    def load_links_from_file(self, site: str, collection_word: str, links_dir: str = "links") -> str:
        """从links目录读取链接数据（委托给FileHandler）"""
        try:
            return self.file_handler.load_links_from_file(site, collection_word, links_dir)
        except Exception as e:
            logger.error(f"从文件加载链接时出错: {str(e)}")
            return ""
    
    def process_site_and_template(self, import_config: dict) -> bool:
        """
        处理站点相关操作（站点选择、导入模板）
        这些操作只需要在站点变化时执行一次
        :param import_config: 导入配置字典
        :return: 是否成功处理
        """
        try:
            site = import_config.get("站点", "")
            if not site:
                print("  ⚠️ 未设置站点，跳过站点相关操作")
                return True  # 没有站点不算失败

            # 1. 选择站点
            print(f"\n🌐 开始选择站点...")
            logger.info(f"选择站点: {site}")
            if not self.select_site(site):
                print(f"  ❌ 站点选择失败")
                return False
            print(f"  ✅ 站点选择完成")
            human_delay(0.5, 0.8)

            # 2. 导入模板（已冻结）
            print(f"\n📁 [已冻结] 跳过导入模板")
            logger.info(f"导入模板功能已冻结，跳过: 站点={site}")

            return True
        except Exception as e:
            logger.error(f"处理站点和模板时出错: {str(e)}")
            print(f"  ❌ 处理站点和模板时出错: {e}")
            traceback.print_exc()
            return False
    
    def process_shop(self, import_config: dict, old_shop: str = "", is_site_changed: bool = False) -> bool:
        """
        处理店铺选择
        只需要在店铺变化时执行
        :param import_config: 导入配置字典
        :param old_shop: 旧店铺名称（用于取消勾选）
        :param is_site_changed: 是否是站点切换（只在站点切换时取消全部勾选）
        :return: 是否成功处理
        """
        try:
            shop_names = import_config.get("店铺", "")
            if not shop_names:
                print("  ⚠️ 未设置店铺，跳过店铺选择")
                return True  # 没有店铺不算失败
            
            # 检查是否是"全部"店铺
            if shop_names.strip() == "全部":
                if is_site_changed:
                    print("\n🏪 站点变化，店铺为'全部'，仅勾选'全部'选项")
                    logger.info("站点变化，店铺为'全部'，仅勾选'全部'选项")
                    if self.select_shops(shop_names, old_shop, is_site_changed, select_all_only=True):
                        print(f"  ✅ '全部'店铺勾选完成")
                        human_delay(0.5, 0.8)
                        return True
                    else:
                        print(f"  ❌ '全部'店铺勾选失败")
                        return False
                else:
                    # 店铺变化时，店铺为"全部"，仅勾选"全部"选项
                    print("\n🏪 店铺变化，店铺为'全部'，仅勾选'全部'选项")
                    logger.info("店铺变化，店铺为'全部'，仅勾选'全部'选项")
                    if self.select_shops(shop_names, old_shop, is_site_changed, select_all_only=True):
                        print(f"  ✅ '全部'店铺勾选完成")
                        human_delay(0.5, 0.8)
                        return True
                    else:
                        print(f"  ❌ '全部'店铺勾选失败")
                        return False

            print(f"\n🏪 开始选择店铺...")
            logger.info(f"选择店铺: {shop_names}, is_site_changed={is_site_changed}")
            if self.select_shops(shop_names, old_shop, is_site_changed):
                print(f"  ✅ 店铺选择完成")
                human_delay(0.5, 0.8)
                return True
            else:
                print(f"  ❌ 店铺选择失败")
                return False
        except Exception as e:
            logger.error(f"处理店铺时出错: {str(e)}")
            print(f"  ❌ 处理店铺时出错: {e}")
            traceback.print_exc()
            return False
    
    def process_single_row_config(self, import_config: dict, row_index: int) -> bool:
        """
        处理单条数据的配置（基础加价、运费、热搜词、链接、类目修改、导入）
        每条数据都需要处理一次
        :param import_config: 导入配置字典
        :param row_index: 数据行索引（用于显示）
        :return: 是否成功处理
        """
        try:
            # 链接始终从links目录读取，不需要从Excel读取
            # 路径格式：links/站点/采集词_*.txt
            site = import_config.get("站点", "")
            collection_word = import_config.get("采集词", "")

            if site and collection_word:
                print(f"\n📂 从links目录读取链接数据...")
                print(f"   站点: {site}")
                print(f"   采集词: {collection_word}")
                print(f"   查找路径: links/{site}/{collection_word}_*.txt")
                logger.info(f"读取链接数据: 站点={site}, 采集词={collection_word}")

                links_data = self.load_links_from_file(site, collection_word)
                if links_data:
                    import_config["链接"] = links_data
                    print(f"  ✅ 链接数据已加载到配置中")
                    print(f"     链接数据长度: {len(links_data)} 字符")
                    link_lines = [l for l in links_data.split('\n') if l.strip()]
                    print(f"     链接行数: {len(link_lines)} 条")
                    logger.info(f"链接数据加载成功，共{len(link_lines)}条")
                else:
                    print(f"  ❌ 未找到对应采集词的链接文件")
                    logger.warning("未找到对应采集词的链接文件")
                    # 记录失败的链接并跳过后续步骤
                    self._save_failed_link(
                        import_config,
                        reason=self._reason_links_file_missing(site, collection_word),
                    )
                    print(f"  ⚠️ 因未找到链接文件，跳过后续步骤")
                    return False
            else:
                print(f"  ⚠️ 缺少站点或采集词信息，无法读取链接文件")
                print(f"     站点: {site if site else '(未设置)'}")
                print(f"     采集词: {collection_word if collection_word else '(未设置)'}")
                logger.warning(f"缺少站点或采集词信息: 站点={site}, 采集词={collection_word}")
                # 如果缺少必要信息，移除链接字段
                import_config.pop("链接", None)

            success_count = 0
            skip_fields = {
                "站点",
                "店铺",
                "分组",
                "采集词",
                "类目",
                "属性",
                "尺寸图",
                "PDF文件",
                "FDApdf文件",
                "TISpdf文件",
                "PS/ICCpdf文件",
            }  # 这些字段不设置到窗口中+只用于内部逻辑

            # 定义字段输入顺序：基础加价 -> 运费 -> 热搜词 -> 重量 -> 链接
            field_order = ["基础加价", "运费", "热搜词", "重量", "链接"]

            # 先按顺序处理定义的字段
            for field_name in field_order:
                if field_name not in import_config:
                    continue

                field_value = import_config[field_name]

                # 跳过"链接"字段中的字典类型（如果是空字典）
                if field_name == "链接" and isinstance(field_value, dict) and not field_value:
                    print(f"  ⚠️ 跳过空字典类型的'链接'字段")
                    continue

                # 处理链接字段
                if field_name == "链接":
                    # 确保链接数据是字符串
                    if isinstance(field_value, str) and field_value.strip():
                        print(f"  ℹ️ 准备设置链接字段，数据长度: {len(field_value)} 字符")
                        if self.set_field_value(field_name, field_value):
                            success_count += 1
                            human_delay(0.5, 0.8)  # 链接数据较大，延迟稍长
                        else:
                            print(f"  ❌ 设置链接字段失败")
                    else:
                        print(f"  ⚠️ 链接字段数据无效或为空")
                    continue

                # 处理热搜词字段：如果为空则先清空再跳过
                if field_name == "热搜词":
                    # 检查值是否为空（None、空字符串、只包含空格、"none"、"None"等）
                    is_empty = False
                    if field_value is None:
                        is_empty = True
                    else:
                        value_str = str(field_value).strip()
                        if not value_str or value_str.lower() == "none":
                            is_empty = True
                    
                    if is_empty:
                        # 先清空热搜词输入框
                        print(f"  ⚠️ 热搜词为空，先清空输入框")
                        self.set_field_value(field_name, "")
                        human_delay(0.3, 0.5)
                        print(f"  ⚠️ 已清空热搜词，跳过填写")
                        continue
                    
                    # 如果值不为空，正常处理
                    value_str = str(field_value).strip()
                    if self.set_field_value(field_name, value_str):
                        success_count += 1
                        human_delay(0.3, 0.5)
                    continue

                # 处理其他字段
                if self.set_field_value(field_name, str(field_value)):
                    success_count += 1
                    human_delay(0.3, 0.5)

            # 处理其他未在顺序列表中的字段（如果有的话）
            for field_name, field_value in import_config.items():
                if field_name in skip_fields or field_name in field_order:
                    continue

                if self.set_field_value(field_name, str(field_value)):
                    success_count += 1
                    human_delay(0.3, 0.5)

            # 填充字段后，将窗口滚动到顶部
            print(f"\n↕️ 将窗口滚动到顶部...")
            if self.scroll_window_to_top():
                print(f"  ✅ 窗口已滚动到顶部")
                logger.info("窗口已成功滚动到顶部")
            else:
                print(f"  ⚠️ 滚动窗口失败，但继续执行")
                logger.warning("滚动窗口到顶部失败")
            human_delay(0.3, 0.5)

            # 类目修改（如果有类目配置）
            category_config = import_config.get("类目", {})
            if category_config and any(category_config.values()):
                print(f"\n📁 开始处理类目修改...")
                logger.info(f"处理类目修改: row_index={row_index - 1}")
                # row_index 是从1开始的，但 process_category_for_row 需要从0开始的索引
                if self.category_automation.process_category_for_row(row_index - 1):
                    print(f"  ✅ 类目修改完成")
                    logger.info("类目修改完成")
                    success_count += 1
                else:
                    print(f"  ❌ 类目修改失败")
                    logger.warning("类目修改失败")
                    # 类目修改失败，记录链接并跳过后续步骤
                    self._save_failed_link(
                        import_config,
                        reason=(
                            getattr(
                                self.category_automation,
                                "_last_category_fail_reason",
                                None,
                            )
                            or "类目修改失败"
                        ),
                    )
                    print(f"  ⚠️ 因类目修改失败，跳过导入按钮点击和后续步骤")
                    return False
            else:
                print(f"\n  ℹ️ 未配置类目信息，跳过类目修改")
                logger.info("未配置类目信息，跳过类目修改")

            human_delay(0.5, 0.8)

            # 点击导入按钮
            print(f"\n📥 开始点击导入按钮...")
            if self.click_import_button(import_config):
                print(f"  ✅ 导入按钮点击完成")
                success_count += 1
            else:
                print(f"  ❌ 导入按钮点击失败")

            return success_count > 0
        except Exception as e:
            logger.error(f"处理单条数据配置时出错: {str(e)}")
            print(f"  ❌ 处理单条数据配置时出错: {e}")
            traceback.print_exc()
            return False
    
    def process_import_config(self, import_config: dict) -> bool:
        """
        处理导入配置
        :param import_config: 导入配置字典
        :return: 是否成功处理所有配置
        """
        try:
            if not import_config:
                logger.warning("未找到导入配置")
                print("\n⚠️ 未找到导入配置")
                return False
            
            # 链接始终从links目录读取，不需要从Excel读取
            # 路径格式：links/站点/采集词_*.txt
            site = import_config.get("站点", "")
            collection_word = import_config.get("采集词", "")
            
            if site and collection_word:
                print(f"\n📂 从links目录读取链接数据...")
                print(f"   站点: {site}")
                print(f"   采集词: {collection_word}")
                print(f"   查找路径: links/{site}/{collection_word}_*.txt")
                logger.info(f"读取链接数据: 站点={site}, 采集词={collection_word}")
                
                links_data = self.load_links_from_file(site, collection_word)
                if links_data:
                    import_config["链接"] = links_data
                    print(f"  ✅ 链接数据已加载到配置中")
                    print(f"     链接数据长度: {len(links_data)} 字符")
                    link_lines = [l for l in links_data.split('\n') if l.strip()]
                    print(f"     链接行数: {len(link_lines)} 条")
                    logger.info(f"链接数据加载成功，共{len(link_lines)}条")
                else:
                    print(f"  ❌ 未找到对应采集词的链接文件")
                    logger.warning("未找到对应采集词的链接文件")
                    # 记录失败的链接并跳过后续步骤
                    self._save_failed_link(
                        import_config,
                        reason=self._reason_links_file_missing(site, collection_word),
                    )
                    print(f"  ⚠️ 因未找到链接文件，跳过后续步骤")
                    return False
            else:
                print(f"  ⚠️ 缺少站点或采集词信息，无法读取链接文件")
                print(f"     站点: {site if site else '(未设置)'}")
                print(f"     采集词: {collection_word if collection_word else '(未设置)'}")
                logger.warning(f"缺少站点或采集词信息: 站点={site}, 采集词={collection_word}")
                # 如果缺少必要信息，移除链接字段
                import_config.pop("链接", None)
            
            print(f"\n📋 开始处理{len(import_config)} 个导入配置项...")
            logger.info(f"开始处理{len(import_config)}个导入配置项")
            
            success_count = 0
            skip_fields = {
                "站点",
                "店铺",
                "分组",
                "尺寸图",
                "PDF文件",
                "FDApdf文件",
                "TISpdf文件",
                "PS/ICCpdf文件",
                "属性",
            }
            
            # 先选择站点（如果有）
            if "站点" in import_config:
                site_name = import_config.get("站点")
                if site_name:
                    print(f"\n🌐 开始选择站点...")
                    logger.info(f"选择站点: {site_name}")
                    if self.select_site(site_name):
                        print(f"  ✅ 站点选择完成")
                        success_count += 1
                    else:
                        print(f"  ❌ 站点选择失败")
                    human_delay(0.5, 0.8)
            
            # 再处理店铺选择（如果有）
            if "店铺" in import_config:
                shop_names = import_config.get("店铺")
                if shop_names:
                    print(f"\n🏪 开始选择店铺...")
                    logger.info(f"选择店铺: {shop_names}")
                    if self.select_shops(shop_names):
                        print(f"  ✅ 店铺选择完成")
                        success_count += 1
                    else:
                        print(f"  ❌ 店铺选择失败")
                    human_delay(0.5, 0.8)
            
            # 选择店铺后，导入模板（如果有站点信息）
            site = import_config.get("站点")
            if site:
                # 导入模板（已冻结）
                print(f"\n📁 [已冻结] 跳过导入模板")
                logger.info(f"导入模板功能已冻结，跳过: 站点={site}")
                
                # 选择模板
                print(f"\n📝 开始选择模板...")
                if self.template_handler.select_template():
                    print(f"  ✅ 模板选择完成")
                    success_count += 1
                else:
                    print(f"  ❌ 模板选择失败")
                human_delay(0.5, 0.8)
            
            # 定义字段输入顺序：基础加价 -> 运费 -> 热搜词 -> 重量 -> 链接
            field_order = ["基础加价", "运费", "热搜词", "重量", "链接"]
            
            # 先按顺序处理定义的字段
            for field_name in field_order:
                if field_name not in import_config:
                    continue
                
                field_value = import_config[field_name]
                
                # 跳过"链接"字段中的字典类型（如果是空字典）
                if field_name == "链接" and isinstance(field_value, dict) and not field_value:
                    print(f"  ⚠️ 跳过空字典类型的'链接'字段")
                    continue
                
                # 处理链接字段
                if field_name == "链接":
                    # 确保链接数据是字符串
                    if isinstance(field_value, str) and field_value.strip():
                        print(f"  ℹ️ 准备设置链接字段，数据长度: {len(field_value)} 字符")
                        if self.set_field_value(field_name, field_value):
                            success_count += 1
                            human_delay(0.5, 0.8)  # 链接数据较大，延迟稍长
                        else:
                            print(f"  ❌ 设置链接字段失败")
                    else:
                        print(f"  ⚠️ 链接字段数据无效或为空")
                    continue
                
                # 处理热搜词字段：如果为空则先清空再跳过
                if field_name == "热搜词":
                    # 检查值是否为空（None、空字符串、只包含空格、"none"、"None"等）
                    is_empty = False
                    if field_value is None:
                        is_empty = True
                    else:
                        value_str = str(field_value).strip()
                        if not value_str or value_str.lower() == "none":
                            is_empty = True
                    
                    if is_empty:
                        # 先清空热搜词输入框
                        print(f"  ⚠️ 热搜词为空，先清空输入框")
                        self.set_field_value(field_name, "")
                        human_delay(0.3, 0.5)
                        print(f"  ⚠️ 已清空热搜词，跳过填写")
                        continue
                    
                    # 如果值不为空，正常处理
                    value_str = str(field_value).strip()
                    if self.set_field_value(field_name, value_str):
                        success_count += 1
                        human_delay(0.3, 0.5)
                    continue
                
                # 处理其他字段
                if self.set_field_value(field_name, str(field_value)):
                    success_count += 1
                    human_delay(0.3, 0.5)
            
            # 处理其他未在顺序列表中的字段（如果有的话）
            for field_name, field_value in import_config.items():
                if field_name in skip_fields or field_name in field_order:
                    continue
                
                if self.set_field_value(field_name, str(field_value)):
                    success_count += 1
                    human_delay(0.3, 0.5)
            
            print(f"\n✅ 成功设置 {success_count}/{len(import_config)} 个字段")
            logger.info(f"字段设置完成: 成功{success_count}/{len(import_config)}个")
            
            # 填充字段后，将窗口滚动到顶部
            print(f"\n↕️ 将窗口滚动到顶部...")
            if self.scroll_window_to_top():
                print(f"  ✅ 窗口已滚动到顶部")
                logger.info("窗口已成功滚动到顶部")
            else:
                print(f"  ⚠️ 滚动窗口失败，但继续执行")
                logger.warning("滚动窗口到顶部失败")
            human_delay(0.3, 0.5)
            
            return success_count > 0
        except Exception as e:
            logger.error(f"处理导入配置时出错: {str(e)}")
            traceback.print_exc()
            return False
    
    def scroll_window_to_top(self, wheel_only: bool = False):
        """向上滚动窗口到最顶部（委托给WindowHandler）"""
        try:
            return self.window_handler.scroll_window_to_top(wheel_only=wheel_only)
        except Exception as e:
            logger.error(f"滚动窗口到顶部时出错: {str(e)}")
            return False
    
    def click_import_template_button(self, site: str, template_dir: str = ".") -> bool:
        """点击导入模板按钮（委托给TemplateHandler）"""
        try:
            return self.template_handler.click_import_template_button(site, template_dir)
        except Exception as e:
            logger.error(f"点击导入模板按钮时出错: {str(e)}")
            return False
    
    def click_size_image_button(self, size_image_file: str) -> bool:
        """点击尺寸图按钮（委托给TemplateHandler）"""
        try:
            return self.template_handler.click_size_image_button(size_image_file)
        except Exception as e:
            logger.error(f"点击尺寸图按钮时出错: {str(e)}")
            return False
    
    def select_template(self) -> bool:
        """选择模板（委托给TemplateHandler）"""
        try:
            return self.template_handler.select_template()
        except Exception as e:
            logger.error(f"选择模板时出错: {str(e)}")
            return False
    
    def scroll_window_to_bottom(self):
        """向下滚动窗口到最底部（委托给WindowHandler）"""
        try:
            return self.window_handler.scroll_window_to_bottom()
        except Exception as e:
            logger.error(f"滚动窗口到底部时出错: {str(e)}")
            return False
    
    def _reason_links_file_missing(self, site: str, collection_word: str) -> str:
        """生成「未找到链接文件」的详细说明（与 FileHandler 查找规则一致）。"""
        links_dir = "links"
        site_dir = os.path.join(links_dir, site or "")
        abs_site = os.path.abspath(site_dir)
        cw = collection_word or ""
        if not site:
            return f"链接文件未找到; 站点字段为空; 采集词={cw!r}"
        if not cw:
            return f"链接文件未找到; 采集词字段为空; 站点={site!r}; 已查目录={abs_site}"
        if not os.path.exists(site_dir):
            return (
                f"链接文件未找到; 站点目录不存在 path={abs_site}; "
                f"采集词={cw}; 期望 {cw}_*.txt 或 {cw}.txt"
            )
        return (
            f"链接文件未找到; 已查目录={abs_site}; 采集词={cw}; "
            f"期望文件名 {cw}_*.txt 或 {cw}.txt（取修改时间最新）"
        )

    def _save_failed_link(self, import_config, reason="未知原因"):
        """保存失败的链接到文件"""
        try:
            from datetime import datetime

            site = import_config.get("站点", "未知站点")
            collection_word = import_config.get("采集词", "未知采集词")
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            reason_safe = (reason or "未知原因").replace("|", "/").replace("\n", " ").strip()

            # 读取已有的失败链接
            existing_links = []
            failed_links_file = "failed_category_links.txt"
            if os.path.exists(failed_links_file):
                try:
                    with open(failed_links_file, "r", encoding="utf-8") as f:
                        existing_links = [line.strip() for line in f.readlines() if line.strip()]
                except Exception:
                    pass

            # 格式: 时间戳|站点|采集词|失败原因（不去重，每次失败追加一行）
            failed_record = f"{timestamp}|{site}|{collection_word}|{reason_safe}"
            all_links = existing_links + [failed_record]

            # 写入文件
            with open(failed_links_file, "w", encoding="utf-8") as f:
                for link_record in all_links:
                    f.write(link_record + "\n")

            print(f"  💾 已将失败记录保存到 {failed_links_file}")
            print(f"     记录: [{timestamp}] 站点={site}, 采集词={collection_word}, 失败原因={reason_safe}")
        except Exception as e:
            print(f"  ⚠️ 保存失败记录时出错: {e}")
    
    def _control_tree_contains_substring(self, ctrl, needle: str, depth: int = 0, max_depth: int = 12) -> bool:
        """判断控件子树中是否出现某段文案（Name / ValuePattern）。"""
        if depth > max_depth:
            return False
        try:
            blob = ctrl.Name or ""
            try:
                vp = ctrl.GetValuePattern()
                if vp and vp.Value:
                    blob = blob + (vp.Value or "")
            except Exception:
                pass
            if needle in blob:
                return True
        except Exception:
            pass
        try:
            child = ctrl.GetFirstChildControl()
            while child:
                if self._control_tree_contains_substring(child, needle, depth + 1, max_depth):
                    return True
                child = child.GetNextSiblingControl()
        except Exception:
            pass
        return False

    def _is_link_error_window_shell(self, w, skip_titles: frozenset) -> bool:
        """是否为可能承载链接校验文案的窗口壳：Win32 #32770 或 WPF 模态 Window。"""
        try:
            if w.ControlType != auto.ControlType.WindowControl:
                return False
            wn = (w.Name or "").strip()
            if wn in skip_titles:
                return False
            w_cls = (w.ClassName or "")
            if w_cls == "#32770":
                return True
            if w_cls == "Window":
                try:
                    wp = w.GetWindowPattern()
                    if wp and wp.IsModal:
                        return True
                except Exception:
                    pass
        except Exception:
            pass
        return False

    def _find_link_error_dialog_in_subtree(self, start_ctrl, markers: tuple, skip_titles: frozenset,
                                           depth: int = 0, max_depth: int = 16):
        """
        在控件子树中查找链接校验失败对话框（含挂在「批量导入」下的 WPF 模态窗）。
        返回 (窗口控件, 匹配到的 marker) 或 None。
        """
        if depth > max_depth:
            return None
        try:
            if self._is_link_error_window_shell(start_ctrl, skip_titles):
                for mk in markers:
                    if self._control_tree_contains_substring(start_ctrl, mk, max_depth=20):
                        return (start_ctrl, mk)
        except Exception:
            pass
        try:
            child = start_ctrl.GetFirstChildControl()
            while child:
                hit = self._find_link_error_dialog_in_subtree(
                    child, markers, skip_titles, depth + 1, max_depth)
                if hit:
                    return hit
                child = child.GetNextSiblingControl()
        except Exception:
            pass
        return None

    def _try_find_win32_prompt_dialog(self):
        """单次探测：标题为「提示」的 Win32 #32770 或顶层模态同名窗。"""
        try:
            d = auto.WindowControl(ClassName="#32770", Name="提示")
            if d.Exists(maxSearchSeconds=0.25):
                return d
        except Exception:
            pass
        try:
            root = auto.GetRootControl()
            for w in root.GetChildren():
                try:
                    if w.ControlType != auto.ControlType.WindowControl:
                        continue
                    if (w.ClassName or "") == "#32770" and (w.Name or "").strip() == "提示":
                        return w
                except Exception:
                    continue
        except Exception:
            pass
        try:
            root = auto.GetRootControl()
            for w in root.GetChildren():
                try:
                    if w.ControlType != auto.ControlType.WindowControl:
                        continue
                    wp = w.GetWindowPattern()
                    if wp and wp.IsModal and (w.Name or "").strip() == "提示":
                        return w
                except Exception:
                    continue
        except Exception:
            pass
        return None

    def _find_link_validation_error_dialog_scan_once(self):
        """
        单次扫描：通过正文识别链接校验失败窗体（Win32 / WPF 模态），无内置 sleep。
        """
        markers = ("多种类型的链接", "不能导入多种类型", "淘宝与拼多多")
        skip_titles = frozenset({"打开", "另存为", "保存"})
        main_hwnd = None
        try:
            if self.window:
                main_hwnd = int(self.window.NativeWindowHandle)
        except Exception:
            pass
        try:
            if self.window:
                hit = self._find_link_error_dialog_in_subtree(
                    self.window, markers, skip_titles)
                if hit:
                    w, mk = hit
                    print(f"  ✓ 在主窗口子树通过正文「{mk}」找到链接校验失败对话框")
                    logger.info(f"主窗口子树找到链接校验失败对话框: {mk}")
                    return w

            root = auto.GetRootControl()
            for w in root.GetChildren():
                try:
                    if w.ControlType != auto.ControlType.WindowControl:
                        continue
                    try:
                        if main_hwnd is not None and int(w.NativeWindowHandle) == main_hwnd:
                            continue
                    except Exception:
                        pass
                    hit = self._find_link_error_dialog_in_subtree(w, markers, skip_titles)
                    if hit:
                        dialog, mk = hit
                        print(f"  ✓ 通过正文「{mk}」找到链接校验失败对话框")
                        logger.info(f"通过正文找到链接校验失败对话框: {mk}")
                        return dialog
                except Exception:
                    continue
        except Exception:
            pass
        return None

    def _click_dialog_confirm_button(self, dialog_window) -> bool:
        """在 Win32 / WPF 对话框上查找并点击「确定」。"""
        confirm_button = None
        try:
            confirm_button = dialog_window.ButtonControl(Name="确定")
            if confirm_button.Exists(maxSearchSeconds=2):
                print("  ✓ 找到'确定'按钮")
                logger.info("找到确定按钮")
            else:
                confirm_button = None
        except Exception:
            confirm_button = None

        if not confirm_button:
            def find_confirm_button(ctrl, depth=0):
                if depth > 5:
                    return None
                try:
                    if ctrl.ControlType == auto.ControlType.ButtonControl:
                        name = (ctrl.Name or "").strip()
                        if name == "确定":
                            return ctrl
                except Exception:
                    pass
                try:
                    child = ctrl.GetFirstChildControl()
                    while child:
                        result = find_confirm_button(child, depth + 1)
                        if result:
                            return result
                        child = child.GetNextSiblingControl()
                except Exception:
                    pass
                return None

            try:
                confirm_button = find_confirm_button(dialog_window)
                if confirm_button:
                    print("  ✓ 通过递归查找找到'确定'按钮")
                    logger.info("通过递归查找找到确定按钮")
            except Exception as e:
                print(f"  ⚠️ 递归查找按钮失败: {e}")

        button_clicked = False
        if confirm_button and confirm_button.Exists(maxSearchSeconds=1):
            try:
                confirm_button.Click()
                print("  ✓ 已通过Click()点击'确定'按钮")
                logger.info("已通过Click()点击确定按钮")
                button_clicked = True
                human_delay(0.5, 0.8)
            except Exception as e:
                print(f"  ⚠️ Click失败: {e}，尝试其他方法...")
                try:
                    invoke_pattern = confirm_button.GetInvokePattern()
                    if invoke_pattern:
                        invoke_pattern.Invoke()
                        print("  ✓ 已通过InvokePattern点击'确定'按钮")
                        logger.info("已通过InvokePattern点击确定按钮")
                        button_clicked = True
                        human_delay(0.5, 0.8)
                except Exception as e2:
                    print(f"  ⚠️ InvokePattern失败: {e2}")
                    if not button_clicked:
                        try:
                            rect = confirm_button.BoundingRectangle
                            if rect.width() > 0 and rect.height() > 0:
                                cx = rect.left + rect.width() // 2
                                cy = rect.top + rect.height() // 2
                                if mouse_click(cx, cy):
                                    print("  ✓ 已通过坐标点击'确定'按钮")
                                    logger.info("已通过坐标点击确定按钮")
                                    button_clicked = True
                                    human_delay(0.5, 0.8)
                        except Exception as e3:
                            print(f"  ⚠️ 坐标点击失败: {e3}")
        return button_clicked

    @staticmethod
    def _filter_to_taoworld_links_only(links_text: str):
        """
        只保留 TaoWorld 标识行（strip 后以 TaoWorld_ 开头，忽略大小写），
        剔除 http(s) 商品链接等。返回 (新文本, 保留行数, 剔除行数)。
        """
        if not links_text or not str(links_text).strip():
            return "", 0, 0
        lines = str(links_text).replace("\r\n", "\n").replace("\r", "\n").split("\n")
        kept = []
        removed = 0
        for line in lines:
            s = line.strip()
            if not s:
                continue
            if s.lower().startswith("taoworld_"):
                kept.append(s)
            else:
                removed += 1
        return "\n".join(kept), len(kept), removed

    def click_import_button(self, import_config=None, *, _link_strip_retry_used=False) -> bool:
        """
        点击导入按钮。
        :param import_config: 当前行配置；链接校验失败时会尝试剔除非 TaoWorld 行并重试一次导入。
        :param _link_strip_retry_used: 内部使用，已做过一次过滤重试后勿再传。
        """
        try:
            if not self.window:
                logger.warning("窗口未初始化，无法点击导入按钮")
                print("❌ 窗口未初始化，无法点击导入按钮")
                return False
            
            print("\n📥 查找并点击'导入'按钮...")
            logger.info("查找并点击导入按钮")
            
            # 先激活窗口
            self.window.SetActive()
            human_delay(0.3, 0.5)
            
            # 先将窗口滚动到最底部
            print(f"\n↕️ 将窗口滚动到最底部...")
            if self.scroll_window_to_bottom():
                print(f"  ✅ 窗口已滚动到最底部")
                logger.info("窗口已成功滚动到最底部")
            else:
                print(f"  ⚠️ 滚动窗口失败，但继续执行")
                logger.warning("滚动窗口到底部失败")
            human_delay(0.3, 0.5)
            
            # 方法1: 优先通过AutomationId查找（最可靠）
            import_button = None
            try:
                import_button = self.window.ButtonControl(AutomationId="AddSkinButton")
                if import_button.Exists(maxSearchSeconds=2):
                    print(f"  ✓ 通过AutomationId找到按钮: AddSkinButton")
                    logger.info("通过AutomationId找到导入按钮: AddSkinButton")
                else:
                    import_button = None
            except:
                import_button = None
            
            # 方法2: 如果AutomationId查找失败，通过按钮名称查找
            if not import_button:
                button_names = ["导入"]
                for name in button_names:
                    try:
                        import_button = self.window.ButtonControl(Name=name)
                        if import_button.Exists(maxSearchSeconds=1):
                            print(f"  ✓ 通过名称找到按钮: {name}")
                            logger.info(f"通过名称找到导入按钮: {name}")
                            break
                        import_button = None
                    except:
                        continue
            
            if not import_button:
                print("  ⚠️ 未找到'导入'按钮")
                logger.warning("未找到导入按钮")
                return False
            
            # 点击按钮 - 使用Click()方法
            try:
                import_button.Click()
                print("  ✓ 已通过Click()点击'导入'按钮")
                logger.info("已通过Click()点击导入按钮")
            except Exception as e:
                print(f"  ❌ 点击'导入'按钮失败: {e}")
                logger.error(f"点击导入按钮失败: {str(e)}")
                return False
            
            # 等待导入后的对话框：「提示」与链接校验失败（WPF）每一轮同时探测，避免先空等「提示」二十多秒
            human_delay(1.0, 1.5)
            print("  🔍 等待导入后的对话框（「提示」或链接校验失败）...")
            logger.info("等待导入后对话框（提示或链接校验）")
            
            try:
                dialog_window = None
                err_dialog = None
                poll_interval = 0.35
                max_wait_sec = 14.0
                attempts = max(8, int(max_wait_sec / poll_interval) + 1)
                for attempt in range(attempts):
                    dialog_window = self._try_find_win32_prompt_dialog()
                    if dialog_window:
                        print(f"  ✓ 找到 Win32「提示」对话框（第 {attempt + 1} 次轮询）")
                        logger.info(f"找到 Win32 提示对话框，尝试 {attempt + 1}")
                        break
                    err_dialog = self._find_link_validation_error_dialog_scan_once()
                    if err_dialog:
                        print(f"  ✓ 第 {attempt + 1} 次轮询检测到链接校验失败弹窗")
                        logger.info(f"检测到链接校验失败弹窗，尝试 {attempt + 1}")
                        break
                    time.sleep(poll_interval)
                
                if err_dialog:
                    print("  ⚠️ 链接校验失败弹窗（混用多平台链接或非 Taoworld 链接等）")
                    logger.warning("检测到链接类型校验失败对话框")
                    link_fail_details = []
                    confirm_closed = self._click_dialog_confirm_button(err_dialog)
                    if confirm_closed:
                        print("  ✓ 已关闭链接校验失败弹窗")
                    else:
                        print("  ⚠️ 未能点击「确定」关闭链接校验失败弹窗")
                        link_fail_details.append("未能自动点击「确定」关闭链接校验弹窗")
                    human_delay(0.4, 0.6)

                    if _link_strip_retry_used:
                        link_fail_details.append("已执行剔除非TaoWorld行并第二次点击导入,仍弹出校验失败")

                    if (
                        import_config is not None
                        and confirm_closed
                        and not _link_strip_retry_used
                    ):
                        raw = import_config.get("链接")
                        if isinstance(raw, str) and raw.strip():
                            filtered, n_kept, n_removed = self._filter_to_taoworld_links_only(raw)
                            if n_removed > 0 and n_kept > 0:
                                import_config["链接"] = filtered
                                print(
                                    f"  🔧 已剔除 {n_removed} 行非 TaoWorld 链接，保留 {n_kept} 行，"
                                    f"正在写回 links 文件并重写链接框、再次点击导入..."
                                )
                                logger.info(
                                    f"链接已过滤: 剔除{n_removed}行, 保留{n_kept}行, 重试导入"
                                )
                                site = (import_config.get("站点") or "").strip()
                                cw = (import_config.get("采集词") or "").strip()
                                if site and cw:
                                    self.file_handler.write_links_to_file(site, cw, filtered)
                                else:
                                    print("  ℹ️ 缺少站点或采集词，跳过写回 links 目录下的 txt")
                                    link_fail_details.append("有可重试过滤结果但缺少站点或采集词,未写回 links 下 txt")
                                if self.set_field_value("链接", filtered):
                                    human_delay(0.5, 0.8)
                                    return self.click_import_button(
                                        import_config,
                                        _link_strip_retry_used=True,
                                    )
                                print("  ❌ 过滤后写回链接字段失败")
                                link_fail_details.append(
                                    f"剔除{n_removed}行非TaoWorld后余{n_kept}行,但写回「链接」输入框失败"
                                )

                            elif n_removed > 0 and n_kept == 0:
                                print("  ⚠️ 过滤后无任何 TaoWorld 链接，无法自动重试")
                                link_fail_details.append(
                                    f"按TaoWorld_前缀过滤后无可保留行(识别{n_removed}行非TaoWorld)"
                                )
                            elif n_removed == 0 and n_kept > 0:
                                link_fail_details.append(
                                    f"共{n_kept}行均以TaoWorld_开头,脚本侧无可剔除项,仍触发应用校验失败"
                                )
                        else:
                            link_fail_details.append(
                                "链接字段非字符串或为空,无法执行剔除非TaoWorld自动重试"
                            )

                    if import_config is not None:
                        base = (
                            "链接校验失败(已剔除非TaoWorld并重试后仍失败)"
                            if _link_strip_retry_used
                            else "链接校验失败(应用侧链接类型/平台校验不通过)"
                        )
                        extra = ("; " + "; ".join(link_fail_details)) if link_fail_details else ""
                        self._save_failed_link(import_config, reason=base + extra)
                    return False
                
                if not dialog_window:
                    print("  ⚠️ 未出现「提示」或链接校验弹窗（可能已关闭或本步无弹窗）")
                    logger.warning("导入后未检测到预期对话框")
                    return True
                
                print("  ✓ 将按「提示」流程处理")
                logger.info("找到提示对话框")
                
                button_clicked = self._click_dialog_confirm_button(dialog_window)
                
                if button_clicked:
                    print("  ✅ 第一个对话框处理完成")
                    logger.info("第一个对话框处理完成")
                    
                    # 等待7秒后按Enter键
                    print("  ⏳ 等待7秒后按Enter键...")
                    logger.info("等待7秒后按Enter键")
                    time.sleep(7.0)
                    
                    try:
                        auto.SendKeys("{ENTER}")
                        print("  ✓ 已按Enter键")
                        logger.info("已按Enter键")
                        human_delay(0.5, 0.8)
                    except Exception as e:
                        print(f"  ⚠️ 按Enter键失败: {e}")
                        logger.warning(f"按Enter键失败: {str(e)}")
                    
                    print("  ✅ 所有对话框处理完成")
                    logger.info("所有对话框处理完成")
                    return True
                else:
                    print("  ⚠️ 未找到'确定'按钮或点击失败")
                    logger.warning("未找到确定按钮或点击失败")
                    # 即使点击失败，也返回True（避免阻塞流程）
                    return True
                    
            except Exception as e:
                print(f"  ⚠️ 处理对话框时出错: {e}")
                logger.warning(f"处理对话框时出错: {str(e)}")
                # 如果处理对话框出错，也算成功（避免阻塞流程）
                return True
                
        except Exception as e:
            logger.error(f"点击导入按钮时出错: {str(e)}")
            print(f"  ❌ 点击导入按钮时出错: {e}")
            traceback.print_exc()
            return False

    def run(self) -> bool:
        """
        执行完整的自动化流程（可处理单个或多个 Excel，每个文件内多条数据按行处理）
        :return: 是否成功完成
        """
        try:
            logger.info("开始批量导入自动化流程")
            print("🎯 开始批量导入自动化流程\n")
            total_start = time.perf_counter()

            # 1. 查找窗口
            if not self.find_window():
                logger.error("查找目标窗口失败，终止流程")
                return False

            excel_paths = list(self.config_manager.excel_file_paths)
            print(f"\n📚 待处理表格文件共 {len(excel_paths)} 个（按顺序）:")
            for i, p in enumerate(excel_paths, 1):
                print(f"   {i}. {p}")
            logger.info(f"Excel 文件列表: {excel_paths}")

            total_success = 0
            total_fail = 0
            total_rows = 0
            template_and_upload_done = False

            for file_num, excel_path in enumerate(excel_paths, 1):
                if not self.config_manager.switch_to_excel(excel_path):
                    logger.error(f"无法切换到 Excel: {excel_path}")
                    continue

                try:
                    all_configs = self.config_manager.get_all_import_configs()
                except Exception as e:
                    logger.error(f"加载配置文件时出错: {str(e)}")
                    all_configs = []

                if not all_configs:
                    print(f"\n⚠️ 跳过文件（无数据）: {excel_path}")
                    logger.warning(f"未找到导入配置，Excel: {excel_path}")
                    continue

                total_rows += len(all_configs)
                print(f"\n{'#'*60}")
                print(f"📂 表格 {file_num}/{len(excel_paths)}: {excel_path}")
                print(f"📋 本文件共 {len(all_configs)} 条数据")
                print(f"{'#'*60}")
                logger.info(
                    f"处理 Excel {file_num}/{len(excel_paths)}: {excel_path}，"
                    f"{len(all_configs)} 条"
                )

                # 第一个「有数据」的表格：一次性模板与上传设置；后续表格仅重置站点/分组等跟踪状态
                if not template_and_upload_done:
                    print(f"\n🔧 执行一次性初始设置...")
                    logger.info("执行一次性初始设置")

                    print(f"\n📝 开始选择模板...")
                    if not self.template_handler.select_template():
                        print(f"  ❌ 模板选择失败")
                        return False
                    print(f"  ✅ 模板选择完成")
                    human_delay(0.5, 0.8)

                    print(f"\n⚙️ 开始配置上传设置...")
                    warehouse_fee = all_configs[0].get("仓储费", "10") if all_configs else "10"
                    self.template_handler.configure_upload_settings(warehouse_fee=warehouse_fee)
                    human_delay(0.5, 0.8)

                    print(f"  ✅ 一次性初始设置完成\n")
                    template_and_upload_done = True

                # 每个新表格重新跟踪站点/分组/店铺/尺寸图，避免沿用上表状态
                current_site = None
                current_group = None
                current_shop = None
                current_size_image = None

                success_count = 0
                fail_count = 0

                for idx, import_config in enumerate(all_configs, 1):
                    # 检查暂停状态
                    self.check_pause()

                    print(f"\n{'='*60}")
                    print(
                        f"📦 [表格 {file_num}/{len(excel_paths)}] "
                        f"第 {idx}/{len(all_configs)} 条数据"
                    )
                    print(f"{'='*60}")
                    logger.info(f"处理第 {idx}/{len(all_configs)} 条数据")

                    # 显示当前数据内容
                    print(f"\n📋 当前数据内容:")
                    for key, value in import_config.items():
                        if key != "链接":  # 链接数据太长，不显示
                            print(f"   {key}: {value}")

                    # 检查站点是否变化
                    site = import_config.get("站点", "")
                    site_changed = False
                    if site and site != current_site:
                        print(f"\n🔄 站点变化: {current_site} -> {site}")
                        logger.info(f"站点变化: {current_site} -> {site}")
                        current_site = site
                        site_changed = True

                        # 检查暂停状态
                        self.check_pause()

                        # 处理站点相关操作（站点选择、导入模板）
                        if not self.process_site_and_template(import_config):
                            print(f"  ❌ 第 {idx} 条数据：站点相关操作失败")
                            fail_count += 1
                            continue

                        # 站点切换后，处理分组选择
                        group = import_config.get("分组", "")
                        if group:
                            print(f"\n🔄 站点切换，选择分组: {group}")
                            logger.info(f"站点切换，选择分组: {group}")
                            self.check_pause()
                            if self.select_group(group):
                                current_group = group
                            else:
                                print(f"  ❌ 第 {idx} 条数据：分组选择失败")
                                fail_count += 1
                                continue

                        # 站点切换时，直接处理店铺选择（不检查店铺是否变化）
                        shop = import_config.get("店铺", "")
                        if shop:
                            print(f"\n🔄 站点切换，检查店铺: {shop}")
                            logger.info(f"站点切换，检查店铺: {shop}")
                        
                            # 检查是否是"全部"店铺
                            if shop.strip() == "全部":
                                print("🏪 站点变化，店铺为'全部'，跳过店铺选择流程")
                                logger.info("站点变化，店铺为'全部'，跳过店铺选择流程")
                            
                                # 更新当前店铺为"全部"，但不执行店铺选择
                                old_shop = current_shop  # 保存旧店铺名称（可能为None）
                                current_shop = shop  # 更新当前店铺
                                # 跳过店铺选择，直接继续处理下一部分
                            else:
                                # 检测到站点切换后，先将滚动条滑到最上面
                                print("📜 检测到站点切换，将滚动条滑到最上面...")
                                self.scroll_window_to_top(wheel_only=True)

                                # 检查暂停状态
                                self.check_pause()

                                # 处理店铺选择（站点切换时，旧店铺信息为空，is_site_changed=True）
                                old_shop = current_shop  # 保存旧店铺名称（可能为None）
                                current_shop = shop  # 更新当前店铺
                                if not self.process_shop(import_config, old_shop, is_site_changed=True):
                                    print(f"  ❌ 第 {idx} 条数据：店铺选择失败")
                                    fail_count += 1
                                    continue
                    else:
                        # 站点没变化时，检查分组是否变化
                        group = import_config.get("分组", "")
                        if group and group != current_group:
                            print(f"\n🔄 分组变化: {current_group} -> {group}")
                            logger.info(f"分组变化: {current_group} -> {group}")
                            self.check_pause()
                            if self.select_group(group):
                                current_group = group
                            else:
                                print(f"  ❌ 第 {idx} 条数据：分组选择失败")
                                fail_count += 1
                                continue

                        # 站点没变化时，才检查店铺是否变化
                        shop = import_config.get("店铺", "")
                        if shop and shop != current_shop:
                            print(f"\n🔄 店铺变化: {current_shop} -> {shop}")
                            logger.info(f"店铺变化: {current_shop} -> {shop}")
                        
                            # 检查是否是"全部"店铺
                            if shop.strip() == "全部":
                                print("🏪 店铺变化，店铺为'全部'，仅勾选'全部'选项")
                                logger.info("店铺变化，店铺为'全部'，仅勾选'全部'选项")
                            
                                # 检查暂停状态
                                self.check_pause()
                            
                                # 检测到店铺变化后，先将滚动条滑到最上面
                                print("📜 检测到店铺变化，将滚动条滑到最上面...")
                                self.scroll_window_to_top(wheel_only=True)
                            
                                # 处理店铺选择（传递旧店铺信息用于取消勾选，is_site_changed=False）
                                old_shop = current_shop  # 保存旧店铺名称
                                current_shop = shop  # 更新当前店铺
                                if not self.process_shop(import_config, old_shop, is_site_changed=False):
                                    print(f"  ❌ 第 {idx} 条数据：店铺选择失败")
                                    fail_count += 1
                                    continue
                            else:
                                # 检查暂停状态
                                self.check_pause()

                                # 检测到店铺变化后，先将滚动条滑到最上面
                                print("📜 检测到店铺变化，将滚动条滑到最上面...")
                                self.scroll_window_to_top(wheel_only=True)

                                # 处理店铺选择（传递旧店铺信息用于取消勾选，is_site_changed=False）
                                old_shop = current_shop  # 保存旧店铺名称
                                current_shop = shop  # 更新当前店铺
                                if not self.process_shop(import_config, old_shop, is_site_changed=False):
                                    print(f"  ❌ 第 {idx} 条数据：店铺选择失败")
                                    fail_count += 1
                                    continue

                    # 检查尺寸图是否变化（独立于站点，根据表格值变化触发）
                    size_image = import_config.get("尺寸图", "")
                    if size_image and size_image != current_size_image:
                        print(f"\n🔄 尺寸图变化: {current_size_image} -> {size_image}")
                        logger.info(f"尺寸图变化: {current_size_image} -> {size_image}")
                        self.check_pause()
                        print(f"\n🖼️ 开始设置尺寸图...")
                        if self.template_handler.click_size_image_button(size_image):
                            print(f"  ✅ 尺寸图设置完成")
                            current_size_image = size_image
                        else:
                            print(f"  ❌ 尺寸图设置失败")
                            fail_count += 1
                            continue
                        human_delay(0.5, 0.8)

                    # 检查暂停状态
                    self.check_pause()

                    # 处理每条数据都需要操作的字段（基础加价、运费、热搜词、链接、类目修改、导入）
                    if self.process_single_row_config(import_config, idx):
                        success_count += 1
                        print(f"  ✅ 第 {idx} 条数据处理完成")
                    else:
                        fail_count += 1
                        print(f"  ❌ 第 {idx} 条数据处理失败")

                total_success += success_count
                total_fail += fail_count
                print(f"\n📊 本表格完成: 成功 {success_count}，失败 {fail_count}（文件: {excel_path}）")
                logger.info(
                    f"表格 {file_num}/{len(excel_paths)} 完成: 成功 {success_count}, 失败 {fail_count}"
                )

            if total_rows == 0:
                print("\n⚠️ 所有表格均无有效数据，未执行导入")
                logger.warning("所有 Excel 均无数据")
                return False

            total_end = time.perf_counter()
            print(f"\n{'='*60}")
            print(f"🏁 全程耗时: {total_end - total_start:.2f} 秒")
            print(
                f"📊 全部表格合计: 成功 {total_success} 条，失败 {total_fail} 条，"
                f"共 {total_rows} 条（{len(excel_paths)} 个文件）"
            )
            print(f"{'='*60}")
            logger.info(
                f"批量导入流程完成，耗时: {total_end - total_start:.2f}秒，"
                f"成功: {total_success}，失败: {total_fail}"
            )

            return total_success > 0
        except Exception as e:
            logger.error(f"执行批量导入流程时出错: {str(e)}")
            traceback.print_exc()
            return False