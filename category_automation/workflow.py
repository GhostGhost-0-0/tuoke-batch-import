"""
类目修改自动化主流程模块
"""

import time
import sys
import os
import uiautomation as auto
import ctypes
import random
from .config import ConfigManager
from .ui_utils import human_delay, mouse_click, get_all_descendants, take_snapshot, safe_click_control
from .dialog_handler import handle_dialog_popup
from .dropdown_handler import select_dropdown_option_with_scroll, select_combobox_by_keyboard
from .attribute_matcher import pair_label_with_input, find_new_controls

# 复用尺寸图的文件对话框逻辑（与尺寸图/PDF 对话框相同）
_parent = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _parent not in sys.path:
    sys.path.insert(0, _parent)
from batch_import_automation.file_dialog_handler import (
    handle_file_select_dialog,
    category_pdf_dialog_title_hint,
)


class CategoryAutomation:
    """类目修改自动化主类"""
    
    WINDOW_NAME = "类目修改"
    
    def __init__(self, config_file: str = "属性配置.json", excel_config_manager=None, pause_event=None):
        """
        初始化自动化类
        :param config_file: 配置文件路径（用于读取属性配置）
        :param excel_config_manager: BatchImportConfigManager实例（用于获取已读取的Excel类目数据）
        :param pause_event: 暂停事件（用于暂停/恢复运行）
        """
        self.config_manager = ConfigManager(config_file)  # 用于读取属性配置
        self.excel_config_manager = excel_config_manager  # 用于获取已读取的Excel类目数据
        self.window = None
        self.current_row_index = 0  # 当前处理的行索引
        self.pause_event = pause_event  # 保存暂停事件
        # 类目流程最后一次失败简述（写入 failed_category_links / 供批量导入复用）
        self._last_category_fail_reason = None

    def set_current_row_index(self, row_index: int):
        """设置当前处理的行索引"""
        self.current_row_index = row_index

    def check_stop(self) -> bool:
        """检查是否需要停止运行"""
        return self.stop_event and self.stop_event.is_set()

    def find_window(self, timeout: int = 5) -> bool:
        """查找并激活目标窗口"""
        self.window = auto.WindowControl(Name=self.WINDOW_NAME, searchDepth=1)
        if not self.window.Exists(maxSearchSeconds=timeout):
            print(f"未找到 '{self.WINDOW_NAME}' 窗口")
            return False
        
        self.window.SetActive()
        human_delay(1.0, 1.5)
        print(f"已找到并激活 '{self.WINDOW_NAME}' 窗口")
        return True
    
    def click_category_modify_button(self) -> bool:
        """
        点击"添加类目"按钮，打开类目修改窗口
        在"批量导入"窗口中查找并点击"添加类目"按钮
        :return: 是否成功点击按钮
        """
        print("\n🔘 查找并点击'添加类目'按钮...")
        
        try:
            # 先查找"批量导入"窗口
            batch_import_window = auto.WindowControl(Name="批量导入", searchDepth=1)
            if not batch_import_window.Exists(maxSearchSeconds=5):
                print("  ❌ 未找到'批量导入'窗口，无法点击添加类目按钮")
                return False
            
            print("  ✓ 找到'批量导入'窗口")
            batch_import_window.SetActive()
            human_delay(0.5, 0.8)
            
            # 方法1: 优先使用AutomationId查找（最可靠）
            button = None
            try:
                button = batch_import_window.ButtonControl(AutomationId="TypeskinButto")
                if button.Exists(maxSearchSeconds=2):
                    print(f"  ✓ 通过AutomationId找到按钮: TypeskinButto")
                else:
                    button = None
            except:
                button = None
            
            if button:
                button_name = button.Name or ""
                print(f"  ✓ 找到按钮: {button_name}")
                
                if safe_click_control(button, timeout=3):
                    print("  ✅ 成功点击'添加类目'按钮")
                    human_delay(1.5, 2.5)  # 等待窗口打开
                    return True
                else:
                    print("  ❌ 点击按钮失败")
                    return False
            else:
                print("  ⚠️ 未找到'添加类目'按钮")
                print("  ℹ️ 提示: 请确保批量导入窗口中存在'添加类目'按钮")
                return False
                
        except Exception as e:
            print(f"  ❌ 查找或点击按钮时出错: {e}")
            return False
    
    def restore_window(self):
        """还原窗口（取消最大化）"""
        try:
            if self.window:
                # 如果窗口已最大化，则还原
                try:
                    window_pattern = self.window.GetWindowPattern()
                    if window_pattern:
                        if window_pattern.WindowVisualState == auto.WindowVisualState.Maximized:
                            window_pattern.SetWindowVisualState(auto.WindowVisualState.Normal)
                            print("已还原窗口（取消最大化）")
                            human_delay(0.5, 1.0)
                except:
                    pass
        except:
            pass
    
    def close_window(self):
        """关闭类目修改窗口，如果失败则停止脚本运行"""
        try:
            if not self.window:
                print("❌ 窗口对象不存在，无法关闭窗口")
                return False
            
            # 首先尝试检查窗口是否已经关闭
            if not self.window.Exists(maxSearchSeconds=1):
                print("✅ 窗口已经关闭")
                return True
            
            # 方法1: 优先使用AutomationId查找关闭按钮（最可靠）
            try:
                close_button = self.window.ButtonControl(AutomationId="ButtonClose")
                if close_button.Exists(maxSearchSeconds=2):
                    print("✓ 找到关闭按钮（通过AutomationId）")
                    try:
                        # 先尝试直接点击
                        close_button.Click()
                        print("已点击关闭按钮（通过AutomationId）")
                        human_delay(1.0, 1.5)  # 增加等待时间
                        
                        # 检查窗口是否已关闭
                        if not self.window.Exists(maxSearchSeconds=1):
                            print("✅ 窗口已成功关闭")
                            return True
                        
                        # 如果窗口未关闭，尝试多次点击
                        for i in range(3):
                            print(f"窗口未关闭，尝试第 {i+1} 次点击...")
                            close_button.Click()
                            human_delay(0.8, 1.2)
                            if not self.window.Exists(maxSearchSeconds=1):
                                print("✅ 窗口已成功关闭")
                                return True
                    except:
                        # 如果点击失败，使用坐标点击
                        r = close_button.BoundingRectangle
                        if r.width() > 0 and r.height() > 0:
                            cx = r.left + r.width() // 2
                            cy = r.top + r.height() // 2
                            print(f"尝试通过坐标点击关闭按钮: ({cx}, {cy})")
                            for i in range(3):
                                if mouse_click(cx, cy):
                                    print(f"已通过坐标点击关闭按钮（第 {i+1} 次）")
                                    human_delay(0.8, 1.2)
                                    if not self.window.Exists(maxSearchSeconds=1):
                                        print("✅ 窗口已成功关闭")
                                        return True
            except Exception as e:
                print(f"使用AutomationId查找关闭按钮时出错: {e}")
            
            # 方法2: 尝试通过Name查找关闭按钮
            try:
                close_button = self.window.ButtonControl(Name="关闭")
                if close_button.Exists(maxSearchSeconds=1):
                    print("✓ 找到关闭按钮（通过Name）")
                    try:
                        close_button.Click()
                        print("已点击关闭按钮（通过Name）")
                        human_delay(1.0, 1.5)
                        
                        # 检查窗口是否已关闭
                        if not self.window.Exists(maxSearchSeconds=1):
                            print("✅ 窗口已成功关闭")
                            return True
                    except:
                        r = close_button.BoundingRectangle
                        if r.width() > 0 and r.height() > 0:
                            cx = r.left + r.width() // 2
                            cy = r.top + r.height() // 2
                            print(f"尝试通过坐标点击关闭按钮: ({cx}, {cy})")
                            for i in range(3):
                                if mouse_click(cx, cy):
                                    print(f"已通过坐标点击关闭按钮（第 {i+1} 次）")
                                    human_delay(0.8, 1.2)
                                    if not self.window.Exists(maxSearchSeconds=1):
                                        print("✅ 窗口已成功关闭")
                                        return True
            except Exception as e:
                print(f"使用Name查找关闭按钮时出错: {e}")
            
            # 方法3: 尝试查找所有按钮控件
            try:
                all_buttons = self.window.GetChildren()
                for button in all_buttons:
                    if button.ControlType == auto.ControlType.ButtonControl:
                        button_name = button.Name or ""
                        if "关闭" in button_name or "Close" in button_name:
                            print(f"✓ 找到可能的关闭按钮: '{button_name}'")
                            try:
                                button.Click()
                                print(f"已点击关闭按钮: '{button_name}'")
                                human_delay(1.0, 1.5)
                                
                                # 检查窗口是否已关闭
                                if not self.window.Exists(maxSearchSeconds=1):
                                    print("✅ 窗口已成功关闭")
                                    return True
                            except:
                                r = button.BoundingRectangle
                                if r.width() > 0 and r.height() > 0:
                                    cx = r.left + r.width() // 2
                                    cy = r.top + r.height() // 2
                                    for i in range(3):
                                        if mouse_click(cx, cy):
                                            print(f"已通过坐标点击关闭按钮（第 {i+1} 次）")
                                            human_delay(0.8, 1.2)
                                            if not self.window.Exists(maxSearchSeconds=1):
                                                print("✅ 窗口已成功关闭")
                                                return True
            except Exception as e:
                print(f"遍历按钮控件时出错: {e}")
            
            # 方法4: 尝试使用Alt+F4
            try:
                self.window.SetFocus()
                human_delay(0.3, 0.5)
                auto.SendKeys("%{F4}")  # Alt+F4
                print("已发送Alt+F4关闭窗口")
                human_delay(1.0, 1.5)
                
                # 检查窗口是否已关闭
                if not self.window.Exists(maxSearchSeconds=1):
                    print("✅ 窗口已成功关闭")
                    return True
            except Exception as e:
                print(f"使用Alt+F4关闭窗口时出错: {e}")
            
            # 方法5: 尝试ESC键
            try:
                self.window.SetFocus()
                human_delay(0.3, 0.5)
                auto.SendKeys("{ESC}")
                print("已发送ESC键关闭窗口")
                human_delay(1.0, 1.5)
                
                # 检查窗口是否已关闭
                if not self.window.Exists(maxSearchSeconds=1):
                    print("✅ 窗口已成功关闭")
                    return True
            except Exception as e:
                print(f"使用ESC键关闭窗口时出错: {e}")
            
            # 方法6: 最后尝试强制关闭
            try:
                print("尝试强制关闭窗口...")
                # 获取窗口句柄
                import win32gui
                import win32con
                
                def find_window_callback(hwnd, windows):
                    if win32gui.IsWindowVisible(hwnd):
                        window_title = win32gui.GetWindowText(hwnd)
                        if "类目修改" in window_title:
                            windows.append(hwnd)
                    return True
                
                windows = []
                win32gui.EnumWindows(find_window_callback, windows)
                
                if windows:
                    for hwnd in windows:
                        try:
                            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                            print("已发送WM_CLOSE消息关闭窗口")
                            human_delay(1.0, 1.5)
                            
                            # 检查窗口是否已关闭
                            if not self.window.Exists(maxSearchSeconds=1):
                                print("✅ 窗口已成功关闭")
                                return True
                        except Exception as e:
                            print(f"发送WM_CLOSE消息时出错: {e}")
            except ImportError:
                print("未安装win32gui，无法尝试强制关闭")
            except Exception as e:
                print(f"强制关闭窗口时出错: {e}")
            
            print("❌ 所有尝试都失败，无法关闭窗口")
            print("请手动关闭类目修改窗口后，重新运行脚本")
            return False
        except Exception as e:
            print(f"❌ 关闭窗口时出错: {e}")
            return False
    
    def is_button_unavailable(self, button) -> bool:
        """
        检测按钮是否不可用（通过LegacyIAccessible.State）
        :param button: 按钮控件
        :return: True表示不可用，False表示可用或无法检测
        """
        try:
            if not button.Exists(maxSearchSeconds=1):
                return False
            
            # 检查是否支持LegacyIAccessiblePattern
            if not button.IsLegacyIAccessiblePatternAvailable:
                # 如果不支持，使用IsEnabled属性作为备用
                return not button.IsEnabled
            
            # 获取LegacyIAccessiblePattern
            legacy_pattern = button.GetLegacyIAccessiblePattern()
            if legacy_pattern:
                # 获取State值（整数）
                state = legacy_pattern.CurrentState
                # 检查是否包含不可用状态 (0x1)
                if state & 0x1:
                    return True
        except Exception as e:
            # 如果获取LegacyIAccessiblePattern失败，使用IsEnabled作为备用
            try:
                return not button.IsEnabled
            except:
                print(f"  检测按钮状态时出错: {e}")
        return False
    
    def wait_for_buttons_available(self, timeout: int = 10) -> bool:
        """
        等待确认按钮和加载类目属性按钮变为可用
        :param timeout: 超时时间（秒）
        :return: 是否成功等待到按钮可用
        """
        if not self.window:
            return False
        
        print("\n⏳ 检测按钮状态，等待按钮变为可用...")
        start_time = time.time()
        
        check_count = 0
        while time.time() - start_time < timeout:
            check_count += 1
            
            # 每次循环都重新查找按钮（因为按钮状态可能变化）
            confirm_button = None
            load_button = None
            
            # 查找确认按钮
            try:
                confirm_button = self.window.ButtonControl(AutomationId="skinButton1")
                if not confirm_button.Exists(maxSearchSeconds=1):
                    confirm_button = self.window.ButtonControl(Name="确认")
                    if not confirm_button.Exists(maxSearchSeconds=1):
                        confirm_button = self.window.ButtonControl(Name="确定")
                        if not confirm_button.Exists(maxSearchSeconds=1):
                            confirm_button = None
            except:
                confirm_button = None
            
            # 查找加载类目属性按钮
            try:
                load_button = self.window.ButtonControl(AutomationId="UpdateSkinButton")
                if not load_button.Exists(maxSearchSeconds=1):
                    load_button = self.window.ButtonControl(Name="加载类目属性")
                    if not load_button.Exists(maxSearchSeconds=1):
                        load_button = None
            except:
                load_button = None
            
            # 如果按钮不存在，跳过本次检测
            if not confirm_button or not load_button:
                if check_count == 1:
                    print("  ⚠️ 未找到按钮，跳过检测")
                time.sleep(0.5)
                continue
            
            # 检测按钮状态
            confirm_unavailable = self.is_button_unavailable(confirm_button)
            load_unavailable = self.is_button_unavailable(load_button)
            
            # 如果两个按钮都可用，返回成功
            if not confirm_unavailable and not load_unavailable:
                elapsed = time.time() - start_time
                print(f"  ✅ 按钮已可用（等待 {elapsed:.1f} 秒，检测 {check_count} 次）")
                return True
            
            # 打印状态（每2秒打印一次）
            if check_count % 4 == 0:
                status = []
                if confirm_unavailable:
                    status.append("确认按钮不可用")
                if load_unavailable:
                    status.append("加载类目属性按钮不可用")
                if status:
                    elapsed = time.time() - start_time
                    print(f"  ⏳ 等待中... ({elapsed:.1f}秒) - {', '.join(status)}")
            
            time.sleep(0.5)
        
        # 超时
        elapsed = time.time() - start_time
        print(f"  ⚠️ 等待按钮可用超时（{elapsed:.1f} 秒）")
        return False
    
    def select_categories(self, category_config: dict) -> bool:
        """选择类目"""
        self._last_category_fail_reason = None
        if not self.window:
            self._last_category_fail_reason = "类目选择失败，未找到类目窗口控件"
            return False

        primary_target = category_config.get("一级", "")
        secondary_target = category_config.get("二级", "")
        tertiary_target = category_config.get("三级", "")
        fourth_target = category_config.get("四级", "")
        fifth_target = category_config.get("五级", "")
        
        primary_combo = self.window.ComboBoxControl(AutomationId="OneTypeSkinComboBox")
        secondary_combo = self.window.ComboBoxControl(AutomationId="TwoTypeSkinComboBox")
        tertiary_combo = self.window.ComboBoxControl(AutomationId="ThreeSkinComboBox")
        fourth_combo = self.window.ComboBoxControl(AutomationId="FourSkinComboBox")
        fifth_combo = self.window.ComboBoxControl(AutomationId="FiveSkinComboBox")
        
        steps = []
        if primary_target:
            steps.append(("一级", primary_combo, primary_target))
        if secondary_target:
            steps.append(("二级", secondary_combo, secondary_target))
        if tertiary_target:
            steps.append(("三级", tertiary_combo, tertiary_target))
        if fourth_target and fourth_combo.Exists(maxSearchSeconds=1):
            steps.append(("四级", fourth_combo, fourth_target))
        if fifth_target and fifth_combo.Exists(maxSearchSeconds=1):
            steps.append(("五级", fifth_combo, fifth_target))
        
        for level_name, combo, target in steps:
            print(f"\n选择{level_name}类目: {target}")
            if not select_combobox_by_keyboard(combo, target):
                self._last_category_fail_reason = (
                    f"类目选择失败，未找到选项: {repr(target)}"
                )
                return False
            human_delay(2.0, 3.0)

        return True
    
    def click_load_non_required_checkbox(self) -> bool:
        """点击'是否加载非必填属性'复选框"""
        if not self.window:
            return False
        
        print("\n🔘 查找并点击'是否加载非必填属性'复选框...")
        
        try:
            # 方法1: 使用AutomationId查找（最可靠）
            checkbox = None
            try:
                checkbox = self.window.CheckBoxControl(AutomationId="chkLoadNonRequired")
                if checkbox.Exists(maxSearchSeconds=2):
                    print(f"  ✓ 通过AutomationId找到复选框: chkLoadNonRequired")
                else:
                    checkbox = None
            except:
                checkbox = None
            
            # 方法2: 如果AutomationId查找失败，使用名称查找
            if not checkbox:
                try:
                    checkbox = self.window.CheckBoxControl(Name="是否加载非必填属性")
                    if checkbox.Exists(maxSearchSeconds=2):
                        print(f"  ✓ 通过名称找到复选框: 是否加载非必填属性")
                    else:
                        checkbox = None
                except:
                    checkbox = None
            
            # 方法3: 如果以上方法都失败，遍历所有控件查找
            if not checkbox:
                def search_checkbox(ctrl, depth=0):
                    if depth > 10:
                        return None
                    try:
                        if ctrl.ControlType == auto.ControlType.CheckBoxControl:
                            name = (ctrl.Name or "").strip()
                            automation_id = (ctrl.AutomationId or "").strip()
                            # 检查名称或AutomationId
                            if name == "是否加载非必填属性" or automation_id == "chkLoadNonRequired":
                                return ctrl
                    except:
                        pass
                    
                    try:
                        child = ctrl.GetFirstChildControl()
                        while child:
                            result = search_checkbox(child, depth + 1)
                            if result:
                                return result
                            child = child.GetNextSiblingControl()
                    except:
                        pass
                    return None
                
                checkbox = search_checkbox(self.window)
            
            if checkbox:
                checkbox_name = checkbox.Name or ""
                print(f"  ✓ 找到复选框: {checkbox_name}")
                
                # 检查复选框状态
                try:
                    toggle_state = checkbox.TogglePattern.ToggleState if hasattr(checkbox, 'TogglePattern') and checkbox.TogglePattern else None
                    if toggle_state == auto.ToggleState.On:
                        print(f"  ℹ️ 复选框已选中，无需点击")
                        return True
                    else:
                        print(f"  ℹ️ 复选框未选中，准备点击")
                except:
                    print(f"  ⚠️ 无法检查复选框状态，直接点击")
                
                # 点击复选框
                if safe_click_control(checkbox, timeout=3):
                    print("  ✅ 成功点击'是否加载非必填属性'复选框")
                    human_delay(0.5, 1.0)  # 等待状态更新
                    return True
                else:
                    print("  ❌ 点击复选框失败")
                    return False
            else:
                print("  ⚠️ 未找到'是否加载非必填属性'复选框")
                # 打印所有复选框名称用于调试
                try:
                    print("  🔍 窗口中的所有复选框:")
                    def print_all_checkboxes(ctrl, depth=0):
                        if depth > 5:
                            return
                        try:
                            if ctrl.ControlType == auto.ControlType.CheckBoxControl:
                                name = (ctrl.Name or "").strip()
                                automation_id = (ctrl.AutomationId or "").strip()
                                print(f"    - 名称: {name}, AutomationId: {automation_id}")
                        except:
                            pass
                        try:
                            child = ctrl.GetFirstChildControl()
                            while child:
                                print_all_checkboxes(child, depth + 1)
                                child = child.GetNextSiblingControl()
                        except:
                            pass
                    print_all_checkboxes(self.window)
                except:
                    pass
                return False
                
        except Exception as e:
            print(f"  ❌ 查找或点击复选框时出错: {e}")
            return False
    
    def load_attributes(self) -> bool:
        """加载属性"""
        if not self.window:
            return False
        
        print("\n拍摄加载前 UI 快照...")
        before_snapshot = take_snapshot(self.window)
        
        # 首先点击"是否加载非必填属性"复选框
        if not self.click_load_non_required_checkbox():
            print("  ⚠️ 未能点击'是否加载非必填属性'复选框，但继续执行")
        
        load_button = self.window.ButtonControl(AutomationId="UpdateSkinButton") or self.window.ButtonControl(Name="加载类目属性")
        if load_button.Exists():
            try:
                load_button.Click()
            except:
                r = load_button.BoundingRectangle
                mouse_click(r.left + r.width() // 2, r.top + r.height() // 2)
            print("已点击'加载类目属性'")
            human_delay(1.0, 1.5)
            
            # 检测弹窗并处理（如果有的话）
            has_popup = handle_dialog_popup(self.window, timeout=3)
            if has_popup:
                print("弹窗已处理，等待窗口稳定...")
                human_delay(0.5, 1.0)
            
            # 点击加载类目属性后，检测按钮状态并等待
            print("\n检测按钮状态，等待按钮变为可用...")
            if not self.wait_for_buttons_available(timeout=10):
                print("  ❌ 等待按钮可用超时，关闭窗口并需要重试")
                self.close_window()
                human_delay(1.0, 1.5)
                return False
            
            # 不在这里点击确认按钮，等属性处理完成后再点击
            print("类目属性已加载，等待处理属性...")
            return True
        else:
            print("未找到'加载类目属性'按钮")
            return False
    
    OPTIONAL_ATTRIBUTES = {
        "FDA化妆品注册编号",
        "TIS标准编号",
        "PS/ICC编号",
    }  # 可选属性，找不到时跳过

    # 多列认证 PDF；找不到上传控件时短重试后跳过，避免长时间卡在滚动与反复遍历 UI 树
    OPTIONAL_PDF_COLUMNS = frozenset({"FDApdf文件", "TISpdf文件", "PS/ICCpdf文件"})

    # Excel 列 -> 用于在界面上定位「认证文件」按钮附近的标签关键词（多 PDF 列）
    _PDF_COLUMN_LABEL_NEEDLES = {
        "FDApdf文件": (
            "FDA化妆品注册",
            "化妆品注册",
            "FDA医疗器械",
            "Thai FDA",
            "Thai FDA cosmetic",
            "cosmetic registration",
            "cosmetic registration number on the product label",
        ),
        "TISpdf文件": (
            "TIS标准编号",
            "TIS certification",
            "Product label with the TIS",
        ),
        "PS/ICCpdf文件": (
            "PS/ICC编号",
            "Philippine Standard (PS)",
            "Import Commodity Clearance (ICC)",
        ),
    }

    def process_attributes(self, attribute_config: dict, pdf_file: str = "", site: str = None, extra_pdf_files: dict = None) -> list:
        """处理属性；先填文本字段，再按列上传 PDF（避免未滚动到对应行时点错按钮）。"""
        if not self.window:
            return []
        
        results = []
        pairs = pair_label_with_input(self.window)
        if not pairs:
            print("\n⚠️ 通过标签匹配未识别到属性字段，将尝试备用方式查找")
        else:
            print(f"\n成功识别 {len(pairs)} 个属性字段")
        for attr_name, attr_value in attribute_config.items():
            print(f"\n处理属性: '{attr_name}' = '{attr_value}'")
            handled = False
            is_optional = attr_name in self.OPTIONAL_ATTRIBUTES
            
            # 特殊处理材质属性
            if attr_name == "材质" or "材质" in attr_name:
                # 尝试通过遍历所有控件查找材质组合框
                material_combo = self.find_material_combobox()
                if material_combo:
                    if select_dropdown_option_with_scroll(material_combo, attr_value, field_name="材质"):
                        handled = True
                        print(f"  已设置材质为: {attr_value}")
                    else:
                        print(f"  设置材质失败: {attr_value}")
                else:
                    print(f"  未找到材质组合框")
            
            # 如果特殊处理失败或不是材质属性，尝试常规处理
            if not handled:
                # 使用更精确的匹配逻辑
                matched = False
                for label, inp in pairs:
                    # 更精确的匹配逻辑：标签名必须包含属性名，或者属性名必须完全匹配标签名
                    # FDA化妆品注册编号 与 TIS标准编号、FDA医疗器械注册号 共用同一 Excel 值，不同类目显示不同控件
                    _label_match = (
                        attr_name in label.Name or label.Name == attr_name
                        or (
                            attr_name == "FDA化妆品注册编号"
                            and ("TIS标准编号" in label.Name or "FDA医疗器械注册号" in label.Name)
                        )
                        or (attr_name == "TIS标准编号" and "TIS标准编号" in label.Name)
                        or (
                            attr_name == "PS/ICC编号"
                            and ("PS/ICC编号" in label.Name or "PS/ICC" in label.Name)
                        )
                    )
                    if _label_match:
                        print(f"  找到匹配的属性字段: '{label.Name}' (属性名: '{attr_name}')")
                        matched = True
                        
                        if inp.ControlType == auto.ControlType.ComboBoxControl:
                            if select_dropdown_option_with_scroll(inp, attr_value, field_name=attr_name):
                                handled = True
                                print(f"  已设置 {attr_name} 为: {attr_value}")
                            else:
                                print(f"  设置 {attr_name} 失败: {attr_value}")
                        elif inp.ControlType == auto.ControlType.EditControl:
                            try:
                                value_pattern = inp.GetValuePattern()
                                if value_pattern:
                                    value_pattern.SetValue(attr_value)
                                    handled = True
                                    print(f"  已设置 {attr_name} 为: {attr_value}")
                                else:
                                    inp.SetFocus()
                                    time.sleep(0.2)
                                    auto.SendKeys("^a")
                                    time.sleep(0.1)
                                    auto.SendKeys(attr_value)
                                    handled = True
                                    print(f"  已通过键盘输入设置 {attr_name} 为: {attr_value}")
                            except Exception as e:
                                print(f"  设置 {attr_name} 失败: {e}")
                        else:
                            print(f"  {attr_name} 字段类型不支持: {inp.ControlType}")
                        
                        break
                
                if not matched:
                    if is_optional:
                        print(f"  标签匹配未找到 '{attr_name}'（可选属性），尝试备用方式...")
                    else:
                        print(f"  标签匹配未找到 '{attr_name}'，尝试备用方式...")
                    try:
                        max_scroll = 3 if is_optional else 15
                        handled = self._find_and_set_attribute_with_scroll(attr_name, attr_value, max_scroll_attempts=max_scroll)
                    except Exception as e:
                        print(f"  ❌ 备用查找出错: {e}")
            
            if not handled:
                if is_optional:
                    print(f"  ℹ️ 此类目无 '{attr_name}' 属性，已跳过")
                    results.append({"name": attr_name, "value": attr_value, "success": True, "skipped": True})
                else:
                    print(f"  ❌ 属性 '{attr_name}' 设置失败")
            else:
                human_delay(0.5, 1.0)
                results.append({"name": attr_name, "value": attr_value, "success": True})

        # PDF：在文本填完后再上传（单列 PDF文件 + FDApdf / TISpdf / PS/ICCpdf）
        extra_pdf_files = extra_pdf_files or {}
        pdf_tasks = []
        if pdf_file and str(pdf_file).strip():
            pdf_tasks.append(("PDF文件", str(pdf_file).strip(), None, None))
        for col_key in ("FDApdf文件", "TISpdf文件", "PS/ICCpdf文件"):
            raw = extra_pdf_files.get(col_key)
            if raw and str(raw).strip():
                needles = self._PDF_COLUMN_LABEL_NEEDLES.get(col_key)
                pdf_tasks.append((col_key, str(raw).strip(), needles, col_key))

        for display_name, path, needles, col_key in pdf_tasks:
            self._try_attach_pdf(path, display_name, site, results, needles, col_key)

        return results
    
    def _find_and_set_attribute_with_scroll(self, attr_name: str, attr_value: str, max_scroll_attempts: int = 15) -> bool:
        """通过遍历窗口查找属性并设置值，支持滚动"""
        user32 = ctypes.windll.user32
        def _find_input():
            all_ctrls = get_all_descendants(self.window)
            for ctrl in all_ctrls:
                if ctrl.ControlType in [auto.ControlType.EditControl, auto.ControlType.ComboBoxControl] and ctrl.IsEnabled:
                    try:
                        child = ctrl.GetFirstChildControl()
                        while child:
                            if child.ControlType == auto.ControlType.TextControl:
                                cn = (child.Name or "").strip()
                                if attr_name in cn or ("FDA" in cn and "FDA" in attr_name) or (
                                    attr_name == "FDA化妆品注册编号"
                                    and ("TIS标准编号" in cn or "FDA医疗器械注册号" in cn)
                                ) or (
                                    attr_name == "TIS标准编号" and "TIS标准编号" in cn
                                ) or (
                                    attr_name == "PS/ICC编号"
                                    and ("PS/ICC" in cn or "PS / ICC" in cn)
                                ):
                                    return ctrl
                            child = child.GetNextSiblingControl()
                    except:
                        pass
            for ctrl in all_ctrls:
                if ctrl.ControlType == auto.ControlType.TextControl:
                    name = (ctrl.Name or "").strip()
                    if attr_name in name or (
                        attr_name == "FDA化妆品注册编号"
                        and ("FDA" in name or "TIS标准编号" in name or "FDA医疗器械注册号" in name)
                    ) or (attr_name == "TIS标准编号" and "TIS标准编号" in name) or (
                        attr_name == "PS/ICC编号" and ("PS/ICC" in name or "PS / ICC" in name)
                    ):
                        label_rect = ctrl.BoundingRectangle
                        best = None
                        min_dx = float('inf')
                        for c in all_ctrls:
                            if c.ControlType in [auto.ControlType.EditControl, auto.ControlType.ComboBoxControl] and c.IsEnabled:
                                r = c.BoundingRectangle
                                if r.width() > 0 and (r.top - 10 < label_rect.bottom and r.bottom + 10 > label_rect.top):
                                    dx = r.left - label_rect.right
                                    if 0 < dx < 350 and dx < min_dx:
                                        min_dx, best = dx, c
                        if best:
                            return best
            return None
        def _set_val(inp):
            if inp.ControlType == auto.ControlType.EditControl:
                try:
                    vp = inp.GetValuePattern()
                    if vp:
                        vp.SetValue(attr_value)
                        return True
                except:
                    pass
                try:
                    inp.SetFocus()
                    time.sleep(0.2)
                    auto.SendKeys("^a")
                    time.sleep(0.1)
                    auto.SendKeys(attr_value)
                    return True
                except:
                    return False
            elif inp.ControlType == auto.ControlType.ComboBoxControl:
                return select_dropdown_option_with_scroll(inp, attr_value, field_name=attr_name)
            return False
        def _visible(c):
            try:
                rect, wr = c.BoundingRectangle, self.window.BoundingRectangle
                return rect.width() > 0 and rect.height() > 0 and wr.top <= rect.top and rect.bottom <= wr.bottom
            except:
                return False
        inp = _find_input()
        if inp and _visible(inp):
            return _set_val(inp)
        win_rect = self.window.BoundingRectangle
        user32.SetCursorPos(int(win_rect.left + win_rect.width() // 2), int(win_rect.top + win_rect.height() // 2))
        time.sleep(0.2)
        for i in range(max_scroll_attempts):
            user32.mouse_event(0x0800, 0, 0, -120, 0)
            time.sleep(0.3)
            inp = _find_input()
            if inp and _visible(inp):
                return _set_val(inp)
        return False

    def _find_pdf_button_near_labels(self, label_needles: tuple, all_ctrls: list = None):
        """根据属性行标签定位同行的「认证文件/PDF」类按钮（多列 PDF 时使用）。"""
        if not self.window or not label_needles:
            return None
        try:
            hints = ("PDF", "认证", "选择", "浏览", "上传", "Certificate", "label", "CPR", "ICC")
            if all_ctrls is None:
                all_ctrls = get_all_descendants(self.window)
            label_rects = []
            for c in all_ctrls:
                if c.ControlType != auto.ControlType.TextControl:
                    continue
                name = (c.Name or "").strip()
                if not name or len(name) > 200:
                    continue
                if any(needle in name for needle in label_needles):
                    label_rects.append(c)
            _btn_types = [auto.ControlType.ButtonControl]
            _hl = getattr(auto.ControlType, "HyperlinkControl", None)
            if _hl is not None:
                _btn_types.append(_hl)
            buttons = [c for c in all_ctrls if c.ControlType in _btn_types]
            best_btn = None
            best_score = float("inf")
            for lab in label_rects:
                try:
                    lr = lab.BoundingRectangle
                except Exception:
                    continue
                for btn in buttons:
                    try:
                        bn = (btn.Name or "").strip()
                        if not bn or not any(h in bn for h in hints):
                            continue
                        br = btn.BoundingRectangle
                        if br.width() <= 0 or br.height() <= 0:
                            continue
                        # 允许按钮在标签右侧或略偏左对齐（同一列布局）
                        if br.right < lr.left - 380:
                            continue
                        if not (br.top - 72 < lr.bottom and br.bottom + 72 > lr.top):
                            continue
                        if br.left >= lr.left - 30:
                            dx = max(0, br.left - lr.right)
                        else:
                            dx = abs(br.left - lr.left) + 8
                        dy = abs((br.top + br.bottom) / 2 - (lr.top + lr.bottom) / 2)
                        score = dx + dy * 0.35
                        if score < best_score:
                            best_score = score
                            best_btn = btn
                    except Exception:
                        continue
            return best_btn
        except Exception as e:
            print(f"  ⚠️ 按标签查找 PDF 按钮出错: {e}")
            return None

    def _try_attach_pdf(self, path: str, display_name: str, site, results: list, label_needles: tuple, column_key: str = None):
        """点击上传按钮；弹窗后在 file_dialog_handler 的长标题列表中查找 #32770（可按列优先匹配标题）。"""
        print(f"\n📄 上传「{display_name}」: {path}")
        dialog_title_hint = None
        if column_key in ("FDApdf文件", "TISpdf文件", "PS/ICCpdf文件"):
            dialog_title_hint = category_pdf_dialog_title_hint(site, column_key)

        optional_pdf = column_key in self.OPTIONAL_PDF_COLUMNS
        max_scroll = 3 if optional_pdf else 18

        pdf_button = None
        if label_needles:
            pdf_button = self._find_pdf_button_near_labels(label_needles)
        if not pdf_button:
            pdf_button = self._find_pdf_button(quick_scan=optional_pdf)
        if not pdf_button and label_needles and self.window:
            print(
                f"  ℹ️ 标签旁未找到按钮，滚动类目窗口后重试（最多 {max_scroll} 次）…",
                flush=True,
            )
            try:
                u32 = ctypes.windll.user32
                wr = self.window.BoundingRectangle
                u32.SetCursorPos(int(wr.left + wr.width() // 2), int(wr.top + wr.height() // 2))
                time.sleep(0.15)
                for _ in range(max_scroll):
                    u32.mouse_event(0x0800, 0, 0, -120, 0)
                    time.sleep(0.22)
                    cached = get_all_descendants(self.window)
                    pdf_button = self._find_pdf_button_near_labels(label_needles, cached)
                    if pdf_button:
                        break
                    pdf_button = self._find_pdf_button(all_ctrls=cached)
                    if pdf_button:
                        break
            except Exception:
                pass
        if not pdf_button:
            if optional_pdf:
                print(f"  ℹ️ 未找到「{display_name}」上传入口（可选认证附件），已跳过")
                results.append(
                    {"name": display_name, "value": path, "success": True, "skipped": True}
                )
            else:
                print(f"  ⚠️ 未找到「{display_name}」对应的选择按钮")
                results.append({"name": display_name, "value": path, "success": False})
            return
        if safe_click_control(pdf_button, timeout=2):
            print(f"  ✓ 已点击「{display_name}」选择按钮")
            human_delay(0.6, 1.0)
            if handle_file_select_dialog(
                path,
                wait_after_click=0,
                site=site,
                dialog_title_hint=dialog_title_hint,
            ):
                print(f"  ✅ 「{display_name}」文件设置完成")
                results.append({"name": display_name, "value": path, "success": True})
            else:
                print(f"  ⚠️ 「{display_name}」文件对话框处理失败")
                results.append({"name": display_name, "value": path, "success": False})
        else:
            print(f"  ⚠️ 点击「{display_name}」按钮失败")
            results.append({"name": display_name, "value": path, "success": False})
    
    def _find_pdf_button(self, all_ctrls=None, quick_scan=False):
        """查找PDF/认证文件选择按钮（all_ctrls 传入时只扫该快照，不执行内置向上滚动，供外层滚动循环复用一次遍历结果）。"""
        if not self.window:
            return None
        keywords = ["选择PDF文件", "PDF", "Thai FDA", "Thai cosmetic", "FDA cosmetic", "认证文件"]
        btn_types = [auto.ControlType.ButtonControl]
        _hl = getattr(auto.ControlType, "HyperlinkControl", None)
        if _hl is not None:
            btn_types.append(_hl)
        btn_types = tuple(btn_types)
        try:
            # 快速路径：常见按钮名，短超时
            for n in ["选择PDF文件", "*Thai FDA cosmetic re...", "选择Thai FDA cosmetic registration number on the product label"]:
                b = self.window.ButtonControl(Name=n)
                if b.Exists(maxSearchSeconds=0.5):
                    return b
            ctrls = all_ctrls if all_ctrls is not None else get_all_descendants(self.window)
            for ctrl in ctrls:
                try:
                    if ctrl.ControlType in btn_types:
                        name = (ctrl.Name or "").strip()
                        if any(k in name for k in keywords):
                            return ctrl
                except Exception:
                    continue
            if all_ctrls is not None:
                return None
            if quick_scan:
                return None
            user32 = ctypes.windll.user32
            wr = self.window.BoundingRectangle
            user32.SetCursorPos(int(wr.left + wr.width() // 2), int(wr.top + wr.height() // 2))
            time.sleep(0.15)
            for _ in range(5):
                user32.mouse_event(0x0800, 0, 0, 120, 0)
                time.sleep(0.2)
                for n in ["选择PDF文件", "*Thai FDA cosmetic re..."]:
                    b = self.window.ButtonControl(Name=n)
                    if b.Exists(maxSearchSeconds=0.5):
                        return b
                for ctrl in get_all_descendants(self.window):
                    try:
                        if ctrl.ControlType in btn_types and any(k in (ctrl.Name or "") for k in keywords):
                            return ctrl
                    except Exception:
                        continue
        except Exception as e:
            print(f"  ⚠️ 查找PDF按钮出错: {e}")
        return None
    
    def find_material_combobox(self):
        """查找材质组合框"""
        if not self.window:
            return None
        
        print("  🔍 查找材质组合框...")
        
        try:
            # 方法1: 先按你确认的 AutomationId 直连查找
            try:
                direct_combo = self.window.ComboBoxControl(AutomationId="selectItem")
                if direct_combo.Exists(maxSearchSeconds=1):
                    try:
                        editable = direct_combo.EditControl(AutomationId="PART_EditableTextBox")
                        if editable.Exists(maxSearchSeconds=0.5):
                            print("    ✅ 通过 AutomationId='selectItem' + PART_EditableTextBox 找到材质组合框")
                            return direct_combo
                    except:
                        pass
                    print("    ✅ 通过 AutomationId='selectItem' 找到材质组合框")
                    return direct_combo
            except:
                pass

            # 方法2: 收集材质文本和 selectItem 组合框，按相对位置匹配
            material_texts = []
            select_item_combos = []
            try:
                for ctrl in get_all_descendants(self.window):
                    try:
                        if ctrl.ControlType == auto.ControlType.TextControl and "材质" in (ctrl.Name or ""):
                            material_texts.append(ctrl)
                            print(f"    找到'材质'文本: '{ctrl.Name}'")
                        elif (
                            ctrl.ControlType == auto.ControlType.ComboBoxControl
                            and (ctrl.AutomationId or "").strip() == "selectItem"
                        ):
                            select_item_combos.append(ctrl)
                    except:
                        continue
            except:
                pass

            if select_item_combos:
                print(f"    找到 {len(select_item_combos)} 个 AutomationId='selectItem' 组合框")
                if material_texts:
                    best_combo = None
                    best_distance = float("inf")
                    for text_ctrl in material_texts:
                        try:
                            text_rect = text_ctrl.BoundingRectangle
                        except:
                            continue
                        for combo in select_item_combos:
                            try:
                                combo_rect = combo.BoundingRectangle
                                if combo_rect.left > text_rect.left and abs(combo_rect.top - text_rect.top) < 80:
                                    distance = combo_rect.left - text_rect.right
                                    if 0 <= distance < best_distance:
                                        best_distance = distance
                                        best_combo = combo
                            except:
                                continue
                    if best_combo:
                        print("    ✅ 通过 selectItem + 标签位置匹配找到材质组合框")
                        return best_combo
                # 没有材质文本或位置匹配失败时，回退到第一个可见组合框
                print("    ✅ 回退使用第一个 AutomationId='selectItem' 组合框")
                return select_item_combos[0]
            
            if not material_texts:
                print("    ⚠️ 未找到'材质'文本")
                return None
            
            # 方法3: 对于每个"材质"文本，查找其右侧的组合框（旧逻辑兜底）
            for material_text in material_texts:
                try:
                    text_rect = material_text.BoundingRectangle
                    print(f"    分析'材质'文本位置: ({text_rect.left}, {text_rect.top}), 大小: {text_rect.width()}x{text_rect.height()}")
                    
                    # 查找所有组合框
                    def find_all_comboboxes(ctrl, depth=0):
                        comboboxes = []
                        if depth > 10:
                            return comboboxes
                        
                        try:
                            if ctrl.ControlType == auto.ControlType.ComboBoxControl:
                                comboboxes.append(ctrl)
                        except:
                            pass
                        
                        # 递归查找子控件
                        try:
                            child = ctrl.GetFirstChildControl()
                            while child:
                                comboboxes.extend(find_all_comboboxes(child, depth + 1))
                                child = child.GetNextSiblingControl()
                        except:
                            pass
                        
                        return comboboxes
                    
                    all_comboboxes = find_all_comboboxes(self.window)
                    print(f"    找到 {len(all_comboboxes)} 个组合框")
                    
                    # 查找最接近且在右侧的组合框
                    best_match = None
                    min_distance = float('inf')
                    
                    for combo in all_comboboxes:
                        try:
                            combo_rect = combo.BoundingRectangle
                            # 检查组合框是否在"材质"文本的右侧且垂直位置相近
                            if (combo_rect.left > text_rect.left and 
                                abs(combo_rect.top - text_rect.top) < 50):
                                # 计算距离
                                distance = combo_rect.left - text_rect.right
                                if distance < min_distance and distance >= 0:
                                    min_distance = distance
                                    best_match = combo
                        except:
                            pass
                    
                    if best_match:
                        try:
                            rect = best_match.BoundingRectangle
                            print(f"    ✅ 找到材质组合框，位置: ({rect.left}, {rect.top}), 距离文本右边: {min_distance}")
                        except:
                            print("    ✅ 找到材质组合框")
                        return best_match
                except Exception as e:
                    print(f"    ❌ 查找组合框时出错: {e}")
                    continue
            
            print("    ⚠️ 未找到匹配的材质组合框")
            
            # 打印所有组合框用于调试
            try:
                print("    🔍 窗口中的所有组合框:")
                def print_all_comboboxes(ctrl, depth=0):
                    if depth > 5:
                        return
                    try:
                        if ctrl.ControlType == auto.ControlType.ComboBoxControl:
                            name = (ctrl.Name or "").strip()
                            automation_id = (ctrl.AutomationId or "").strip()
                            try:
                                rect = ctrl.BoundingRectangle
                                print(f"      - 名称: '{name}', AutomationId: '{automation_id}', 位置: ({rect.left}, {rect.top}), 大小: {rect.width()}x{rect.height()}")
                            except:
                                print(f"      - 名称: '{name}', AutomationId: '{automation_id}'")
                    except:
                        pass
                    try:
                        child = ctrl.GetFirstChildControl()
                        while child:
                            print_all_comboboxes(child, depth + 1)
                            child = child.GetNextSiblingControl()
                    except:
                        pass
                print_all_comboboxes(self.window)
                
                # 打印所有文本控件用于调试
                print("    🔍 窗口中的所有文本控件:")
                def print_all_texts(ctrl, depth=0):
                    if depth > 5:
                        return
                    try:
                        if ctrl.ControlType == auto.ControlType.TextControl:
                            name = (ctrl.Name or "").strip()
                            try:
                                rect = ctrl.BoundingRectangle
                                print(f"      - 文本: '{name}', 位置: ({rect.left}, {rect.top}), 大小: {rect.width()}x{rect.height()}")
                            except:
                                print(f"      - 文本: '{name}'")
                    except:
                        pass
                    try:
                        child = ctrl.GetFirstChildControl()
                        while child:
                            print_all_texts(child, depth + 1)
                            child = child.GetNextSiblingControl()
                    except:
                        pass
                print_all_texts(self.window)
            except:
                pass
            
            return None
                
        except Exception as e:
            print(f"    ❌ 查找材质组合框时出错: {e}")
            return None
    
    def click_confirm(self):
        """点击确认按钮"""
        if not self.window:
            return False
        
        try:
            # 优先使用 AutomationId 查找（更可靠）
            confirm_button = self.window.ButtonControl(AutomationId="skinButton1")
            if not confirm_button.Exists(maxSearchSeconds=2):
                # 如果没找到，尝试通过 Name 查找
                confirm_button = self.window.ButtonControl(Name="确认")
                if not confirm_button.Exists(maxSearchSeconds=2):
                    # 再尝试"确定"
                    confirm_button = self.window.ButtonControl(Name="确定")
                    if not confirm_button.Exists(maxSearchSeconds=2):
                        print("未找到确认按钮")
                        return False
            
            # 统一使用鼠标移动点击按钮
            try:
                r = confirm_button.BoundingRectangle
                if r.width() > 0 and r.height() > 0:
                    cx = r.left + r.width() // 2
                    cy = r.top + r.height() // 2
                    if mouse_click(cx, cy):
                        print("已通过鼠标移动点击确认按钮")
                        human_delay(0.5, 1.0)
                        return True
                    else:
                        print("鼠标移动点击确认按钮失败")
                        return False
                else:
                    print("确认按钮尺寸无效")
                    return False
            except Exception as e:
                print(f"鼠标移动点击确认按钮时出错: {e}")
                return False
        except Exception as e:
            print(f"点击确认按钮时出错: {e}")
            return False
    
    def save_results(self, results: list):
        """保存结果"""
        if results:
            success_count = sum(1 for r in results if r.get('success'))
            total_count = len(results)
            print(f"\n🎯 属性处理结果: 成功 {success_count}/{total_count} 个属性")
            
            # 详细显示每个属性的处理结果
            for result in results:
                status = "✅ 成功" if result.get('success') else "❌ 失败"
                print(f"    {status}: {result.get('name', '未知属性')} = {result.get('value', '未知值')}")
        else:
            print("\nℹ️ 没有属性需要处理或处理结果为空")
    
    def run(self, row_index=0) -> bool:
        """执行完整的自动化流程"""
        print("开始类目修改自动化流程\n")
        total_start = time.perf_counter()
        
        # 1. 点击"添加类目"按钮，打开类目修改窗口
        if not self.click_category_modify_button():
            print("❌ 点击'添加类目'按钮失败，终止流程")
            return False
        
        # 2. 查找窗口
        if not self.find_window():
            return False
        
        # 3. 取消最大化窗口（如果已最大化，则还原）
        self.restore_window()
        
        # 4. 加载配置
        # 类目配置从Excel读取（使用已读取的数据）
        if self.excel_config_manager:
            category_config = self.excel_config_manager.get_category_config(row_index)
            print(f"\n📋 从Excel读取类目配置:")
            for level, value in category_config.items():
                if value:
                    print(f"   {level}类目: {value}")
        else:
            # 如果没有提供excel_config_manager，从JSON读取（兼容旧方式）
            config = self.config_manager.load()
            category_config = config.get("类目", {})
            print(f"\n⚠️ 未提供Excel配置管理器，从JSON读取类目配置")
        
        # 属性配置读取
        if self.excel_config_manager:
            # 从Excel读取属性配置
            config = self.excel_config_manager.get_import_config(row_index)
            attribute_config = config.get("属性", {})
            
            # 过滤掉空值
            attribute_config = {k: v for k, v in attribute_config.items() if v is not None and v != ""}
            
            if attribute_config:
                print(f"\n📋 从Excel读取属性配置:")
                for attr_name, attr_value in attribute_config.items():
                    print(f"   {attr_name}: {attr_value}")
        else:
            # 兼容旧方式，从JSON读取属性配置
            config = self.config_manager.load()
            attribute_config = config.get("属性", {})
        
        # 5. 选择类目
        if not self.select_categories(category_config):
            print("类目选择失败")
            return False
        
        # 6. 加载属性
        if not self.load_attributes():
            print("加载属性失败")
            return False
        
        # 7. 处理属性
        results = []
        if not self.excel_config_manager:
            config = self.config_manager.load()
        pdf_file = (config or {}).get("PDF文件", "") or ""
        site = (config or {}).get("站点", "") or ""
        extra_pdf = {
            k: (config.get(k) or "").strip()
            for k in ("FDApdf文件", "TISpdf文件", "PS/ICCpdf文件")
            if config and (config.get(k) or "").strip()
        }
        if attribute_config or pdf_file or extra_pdf:
            results = self.process_attributes(
                attribute_config or {},
                pdf_file=pdf_file,
                site=site,
                extra_pdf_files=extra_pdf or None,
            )
        
        # 8. 点击确认按钮
        if results or not attribute_config or pdf_file or extra_pdf:
            print("点击类目修改窗口的确认按钮...")
            if self.click_confirm():
                print("  ✅ 确认按钮点击成功")
            else:
                print("  ❌ 确认按钮点击失败，但继续执行")
        else:
            print("  ⚠️ 属性处理失败，不点击确认按钮")
        
        # 9. 保存结果
        self.save_results(results)
        
        total_end = time.perf_counter()
        print(f"\n全程耗时: {total_end - total_start:.2f} 秒")
        print("类目修改配置设置完成！")
        
        return True
    
    def process_category_for_row(self, row_index: int) -> bool:
        """
        处理指定行的类目配置
        :param row_index: 行索引
        :return: 是否成功处理
        """
        max_retries = 3  # 最多重试2次（总共3次尝试）
        failed_links = []  # 记录失败的链接
        self._last_category_fail_reason = None

        for attempt in range(max_retries):
            if attempt > 0:
                print(f"\n🔄 第 {row_index + 1} 条数据：重试处理（第 {attempt + 1} 次尝试）...")
            
            print(f"\n🔘 处理第 {row_index + 1} 条数据的类目配置...")
            
            # 1. 点击"添加类目"按钮，打开类目修改窗口
            if not self.click_category_modify_button():
                print(f"  ❌ 第 {row_index + 1} 条数据：点击'添加类目'按钮失败")
                if attempt < max_retries - 1:
                    # 重试前等待一段时间
                    human_delay(1.0, 2.0)
                    continue  # 重试
                else:
                    # 最后一次尝试失败，记录链接并跳过
                    if self.excel_config_manager:
                        config = self.excel_config_manager.get_import_config(row_index)
                        self._last_category_fail_reason = (
                            "添加类目失败，无法点击「添加类目」按钮"
                        )
                        failed_link = {
                            "site": config.get("站点", "未知站点"),
                            "collection_word": config.get("采集词", "未知采集词"),
                            "link": config.get("链接", "未知链接"),
                            "reason": self._last_category_fail_reason,
                        }
                        failed_links.append(failed_link)
                        self._save_failed_links(failed_links)
                    print(f"  ❌ 第 {row_index + 1} 条数据：点击'添加类目'按钮失败，已达到最大重试次数，跳过该条数据")
                    return False
            
            # 2. 查找窗口
            if not self.find_window():
                print(f"  ❌ 第 {row_index + 1} 条数据：查找类目修改窗口失败")
                if attempt < max_retries - 1:
                    # 重试前等待一段时间
                    human_delay(1.0, 2.0)
                    continue  # 重试
                else:
                    # 最后一次尝试失败，记录链接并跳过
                    if self.excel_config_manager:
                        config = self.excel_config_manager.get_import_config(row_index)
                        self._last_category_fail_reason = (
                            f"类目窗口失败，未找到「{self.WINDOW_NAME}」"
                        )
                        failed_link = {
                            "site": config.get("站点", "未知站点"),
                            "collection_word": config.get("采集词", "未知采集词"),
                            "link": config.get("链接", "未知链接"),
                            "reason": self._last_category_fail_reason,
                        }
                        failed_links.append(failed_link)
                        self._save_failed_links(failed_links)
                    print(f"  ❌ 第 {row_index + 1} 条数据：查找类目修改窗口失败，已达到最大重试次数，跳过该条数据")
                    return False
            
            # 3. 取消最大化窗口（如果已最大化，则还原）
            self.restore_window()
            
            # 4. 加载指定行的类目配置
            if self.excel_config_manager:
                category_config = self.excel_config_manager.get_category_config(row_index)
                print(f"\n📋 第 {row_index + 1} 条数据的类目配置:")
                for level, value in category_config.items():
                    if value:
                        print(f"   {level}类目: {value}")
                
                # 如果类目配置为空，跳过处理
                if not any(category_config.values()):
                    print(f"  ⚠️ 第 {row_index + 1} 条数据：类目配置为空，跳过类目处理")
                    return True
            else:
                # 如果没有提供excel_config_manager，从JSON读取（兼容旧方式）
                config = self.config_manager.load()
                category_config = config.get("类目", {})
                print(f"\n⚠️ 第 {row_index + 1} 条数据：未提供Excel配置管理器，从JSON读取类目配置")
            
            # 5. 选择类目
            if not self.select_categories(category_config):
                print(f"  ❌ 第 {row_index + 1} 条数据：类目选择失败")
                if attempt < max_retries - 1:
                    # 关闭窗口并重试
                    try:
                        self.close_window()
                        human_delay(1.0, 2.0)
                    except:
                        pass
                    continue  # 重试
                else:
                    # 最后一次尝试失败，记录链接并跳过
                    try:
                        # 尝试关闭窗口
                        self.close_window()
                    except:
                        pass
                    if self.excel_config_manager:
                        config = self.excel_config_manager.get_import_config(row_index)
                        self._last_category_fail_reason = (
                            self._last_category_fail_reason or "类目选择失败"
                        )
                        failed_link = {
                            "site": config.get("站点", "未知站点"),
                            "collection_word": config.get("采集词", "未知采集词"),
                            "link": config.get("链接", "未知链接"),
                            "reason": self._last_category_fail_reason,
                        }
                        failed_links.append(failed_link)
                        self._save_failed_links(failed_links)
                    print(f"  ❌ 第 {row_index + 1} 条数据：类目选择失败，已达到最大重试次数，跳过该条数据")
                    return False
            
            # 6. 加载属性
            if not self.load_attributes():
                print(f"  ❌ 第 {row_index + 1} 条数据：加载属性失败（按钮一直不可用）")
                if attempt < max_retries - 1:
                    # 窗口已在load_attributes中关闭，等待一下后重试
                    human_delay(1.0, 2.0)
                    continue  # 重试
                else:
                    # 最后一次尝试失败，记录链接并跳过
                    if self.excel_config_manager:
                        config = self.excel_config_manager.get_import_config(row_index)
                        self._last_category_fail_reason = (
                            "加载属性失败，等待控件超时或未找到「加载类目属性」"
                        )
                        failed_link = {
                            "site": config.get("站点", "未知站点"),
                            "collection_word": config.get("采集词", "未知采集词"),
                            "link": config.get("链接", "未知链接"),
                            "reason": self._last_category_fail_reason,
                        }
                        failed_links.append(failed_link)
                        self._save_failed_links(failed_links)
                    print(f"  ❌ 第 {row_index + 1} 条数据：加载属性失败，已达到最大重试次数，跳过该条数据")
                    return False
            
            # 7. 处理属性
            results = []
            if self.excel_config_manager:
                # 从Excel读取属性配置
                config = self.excel_config_manager.get_import_config(row_index)
                attribute_config = config.get("属性", {})
                
                # 过滤掉空值
                attribute_config = {k: v for k, v in attribute_config.items() if v is not None and v != ""}
                
                if attribute_config:
                    print(f"\n📋 第 {row_index + 1} 条数据的属性配置:")
                    for attr_name, attr_value in attribute_config.items():
                        print(f"   {attr_name}: {attr_value}")
                if config.get("PDF文件"):
                    print(f"   PDF文件: {config.get('PDF文件')}")
                for _pk in ("FDApdf文件", "TISpdf文件", "PS/ICCpdf文件"):
                    if config.get(_pk):
                        print(f"   {_pk}: {config.get(_pk)}")
                pdf_file = config.get("PDF文件", "") or ""
                site = config.get("站点", "") or ""
                extra_pdf = {
                    k: (config.get(k) or "").strip()
                    for k in ("FDApdf文件", "TISpdf文件", "PS/ICCpdf文件")
                    if (config.get(k) or "").strip()
                }
                if attribute_config or pdf_file or extra_pdf:
                    results = self.process_attributes(
                        attribute_config or {},
                        pdf_file=pdf_file,
                        site=site,
                        extra_pdf_files=extra_pdf or None,
                    )
                    if results:
                        print(f"  ✅ 属性处理完成，共处理 {len(results)} 个属性")
                    else:
                        print(f"  ⚠️ 属性处理未返回结果")
                else:
                    print(f"  ℹ️ 没有配置属性需要处理")
            else:
                # 兼容旧方式，从JSON读取属性配置
                config = self.config_manager.load()
                attribute_config = config.get("属性", {})
                attribute_config = {k: v for k, v in attribute_config.items() if v is not None and v != ""}
                if attribute_config:
                    print(f"\n📋 从JSON读取属性配置:")
                    for attr_name, attr_value in attribute_config.items():
                        print(f"   {attr_name}: {attr_value}")
                pdf_file = (config or {}).get("PDF文件", "") or ""
                site = (config or {}).get("站点", "") or ""
                extra_pdf = {
                    k: ((config or {}).get(k) or "").strip()
                    for k in ("FDApdf文件", "TISpdf文件", "PS/ICCpdf文件")
                    if ((config or {}).get(k) or "").strip()
                }
                if attribute_config or pdf_file or extra_pdf:
                    results = self.process_attributes(
                        attribute_config or {},
                        pdf_file=pdf_file,
                        site=site,
                        extra_pdf_files=extra_pdf or None,
                    )
                    if results:
                        print(f"  ✅ 属性处理完成，共处理 {len(results)} 个属性")
                    else:
                        print(f"  ⚠️ 属性处理未返回结果")
                else:
                    print(f"  ℹ️ JSON中没有配置属性")
            
            # 8. 点击确认按钮
            if results or (not attribute_config and not pdf_file and not extra_pdf):
                print("点击类目修改窗口的确认按钮...")
                if self.click_confirm():
                    print("  ✅ 确认按钮点击成功")
                else:
                    print("  ❌ 确认按钮点击失败，但继续执行")
            else:
                print("  ⚠️ 属性处理失败，不点击确认按钮")
            
            # 9. 保存结果
            self.save_results(results)
            
            print(f"  ✅ 第 {row_index + 1} 条数据：类目修改配置设置完成")
            return True
        
        # 所有重试都失败
        print(f"  ❌ 第 {row_index + 1} 条数据：经过 {max_retries} 次尝试后仍然失败")
        return False
    
    def _save_failed_links(self, failed_links):
        """保存失败的链接到文件（行格式: 时间戳|站点|采集词|原因；不去重，追加）。"""
        from datetime import datetime

        try:
            if not failed_links:
                return

            # 读取已有的失败链接
            existing_links = []
            failed_links_file = "failed_category_links.txt"
            if os.path.exists(failed_links_file):
                try:
                    with open(failed_links_file, "r", encoding="utf-8") as f:
                        existing_links = [line.strip() for line in f.readlines() if line.strip()]
                except Exception:
                    pass

            # 处理新失败的链接
            processed_links = []
            for link in failed_links:
                if isinstance(link, str):
                    processed_links.append(link)
                elif isinstance(link, dict):
                    site = link.get("site", "未知站点")
                    collection_word = link.get("collection_word", "未知采集词")
                    reason = link.get("reason", "未知原因")
                    reason_safe = (reason or "未知原因").replace("|", "/").replace("\n", " ").strip()
                    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    record = f"{ts}|{site}|{collection_word}|{reason_safe}"
                    processed_links.append(record)
                    print(f"     失败原因: {reason_safe}")

            all_links = existing_links + processed_links

            with open(failed_links_file, "w", encoding="utf-8") as f:
                for link_record in all_links:
                    f.write(link_record + "\n")

            print(f"  💾 已将 {len(failed_links)} 个失败的记录保存到 {failed_links_file}")
        except Exception as e:
            print(f"  ⚠️ 保存失败记录时出错: {e}")

