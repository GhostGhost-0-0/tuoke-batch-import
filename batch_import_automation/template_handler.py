"""
模板操作处理模块
包含：导入模板、尺寸图设置、选择模板等功能
"""

import time
import ctypes

import uiautomation as auto

from .ui_utils import human_delay, mouse_click
from .file_dialog_handler import handle_file_select_dialog

# Windows API 用于鼠标事件
user32 = ctypes.windll.user32

# 选择模板：跨进程嵌套 UIA 下 pywinauto 全树 descendants 可能极慢或长时间卡住，这里用有界深度的 uia 遍历
_TEMPLATE_UI_MAX_DEPTH = 22


def _iter_uia_depth_first(root: auto.Control, max_depth: int):
    """深度优先遍历控件，限制最大深度，避免扫全桌面 UIA 树。"""

    def walk(ctrl: auto.Control, depth: int):
        if depth > max_depth:
            return
        yield ctrl
        if depth == max_depth:
            return
        try:
            child = ctrl.GetFirstChildControl()
        except Exception:
            return
        while child:
            yield from walk(child, depth + 1)
            try:
                child = child.GetNextSiblingControl()
            except Exception:
                break

    yield from walk(root, 0)


def _first_combo_under(root: auto.Control):
    for c in _iter_uia_depth_first(root, _TEMPLATE_UI_MAX_DEPTH):
        try:
            if c.ControlType == auto.ControlType.ComboBoxControl:
                return c
        except Exception:
            continue
    return None


def _list_item_text_matches(item: auto.Control, target: str) -> bool:
    try:
        if (item.Name or "").strip() == target:
            return True
    except Exception:
        pass
    try:
        ch = item.GetFirstChildControl()
        while ch:
            if ch.ControlType == auto.ControlType.TextControl:
                if (ch.Name or "").strip() == target:
                    return True
            ch = ch.GetNextSiblingControl()
    except Exception:
        pass
    return False


def _first_list_item_named(root: auto.Control, target: str, max_depth=None):
    md = _TEMPLATE_UI_MAX_DEPTH if max_depth is None else max_depth
    for c in _iter_uia_depth_first(root, md):
        try:
            if c.ControlType != auto.ControlType.ListItemControl:
                continue
            if _list_item_text_matches(c, target):
                return c
        except Exception:
            continue
    return None


def _first_list_item_named_under_top_windows(target: str, max_depth=14, time_budget_sec=5.0):
    """下拉列表偶发挂在独立顶层浮层上：在预算内浅扫各顶层窗口。"""
    deadline = time.monotonic() + time_budget_sec
    try:
        root = auto.GetRootControl()
        child = root.GetFirstChildControl()
    except Exception:
        return None
    while child and time.monotonic() < deadline:
        try:
            if child.ControlType == auto.ControlType.WindowControl:
                hit = _first_list_item_named(child, target, max_depth=max_depth)
                if hit:
                    return hit
        except Exception:
            pass
        try:
            child = child.GetNextSiblingControl()
        except Exception:
            break
    return None


def _wait_template_edit_window(timeout_sec: float = 45.0, poll: float = 0.45):
    """轮询顶层「模板编辑」窗口；超时失败而非无限阻塞。"""
    deadline = time.monotonic() + timeout_sec
    shown_wait_log = False
    while time.monotonic() < deadline:
        try:
            w = auto.WindowControl(Name="模板编辑", searchDepth=1)
            if w.Exists(maxSearchSeconds=min(0.5, poll)):
                return w
        except Exception:
            pass
        if not shown_wait_log:
            print("  ⏳ 等待「模板编辑」窗口出现（若跨进程 UIA 较慢，会比平时久一点）…", flush=True)
            shown_wait_log = True
        time.sleep(poll)
    return None


class TemplateHandler:
    """模板操作处理器"""
    
    def __init__(self, window=None):
        """
        初始化模板处理器
        :param window: 主窗口控件
        """
        self.window = window
    
    def click_import_template_button(self, site: str, template_dir: str = ".") -> bool:
        """
        点击"导入模板"按钮，并在文件选择对话框中选择对应的wpfTpl文件
        :param site: 站点名称（如"台湾"、"新加坡"等）
        :param template_dir: 模板文件所在目录
        :return: 是否成功完成导入模板操作
        """
        if not self.window:
            print("❌ 窗口未初始化，无法点击导入模板按钮")
            return False
        
        print(f"\n📁 开始导入模板操作...")
        print(f"   站点: {site}")
        
        try:
            # 1. 查找并点击"导入模板"按钮
            import_button = self.window.ButtonControl(AutomationId="ImportTemplage")
            if not import_button.Exists(maxSearchSeconds=2):
                print("  ⚠️ 未找到'导入模板'按钮")
                return False
            
            print("  ✓ 找到'导入模板'按钮")
            
            # 点击按钮
            try:
                import_button.Click()
                print("  ✓ 已点击'导入模板'按钮")
            except:
                # 如果Click失败，尝试使用InvokePattern
                try:
                    invoke_pattern = import_button.GetInvokePattern()
                    if invoke_pattern:
                        invoke_pattern.Invoke()
                        print("  ✓ 已通过InvokePattern点击'导入模板'按钮")
                except:
                    print("  ❌ 点击'导入模板'按钮失败")
                    return False
            
            human_delay(1.5, 2.0)  # 等待对话框出现
            
            # 2. 查找文件选择对话框（Win32对话框）
            print("  🔍 等待文件选择对话框出现...")
            dialog = None
            
            # 方法1: 直接通过ClassName和Name查找（最直接）
            for attempt in range(20):  # 最多等待10秒
                try:
                    dialog = auto.WindowControl(ClassName="#32770", Name="打开")
                    if dialog.Exists(maxSearchSeconds=0.5):
                        print(f"  ✓ 找到Win32文件选择对话框（通过ClassName和Name），尝试次数: {attempt + 1}")
                        break
                    dialog = None
                except:
                    pass
                time.sleep(0.5)
            
            # 方法2: 如果方法1失败，遍历所有窗口查找（添加调试信息）
            if not dialog:
                print("  🔍 方法1未找到，尝试遍历所有窗口...")
                for attempt in range(20):  # 最多等待10秒
                    try:
                        root = auto.GetRootControl()
                        all_windows = root.GetChildren()
                        
                        # 打印所有窗口信息用于调试（仅第一次）
                        if attempt == 0:
                            print("  📋 当前所有窗口列表:")
                            for idx, w in enumerate(all_windows[:10]):  # 只显示前10个
                                try:
                                    if w.ControlType == auto.ControlType.WindowControl:
                                        w_name = (w.Name or "").strip()
                                        w_class = w.ClassName or ""
                                        print(f"    {idx+1}. Name='{w_name}', ClassName='{w_class}'")
                                except:
                                    pass
                        
                        for w in all_windows:
                            try:
                                if w.ControlType == auto.ControlType.WindowControl:
                                    w_name = (w.Name or "").strip()
                                    w_class = w.ClassName or ""
                                    
                                    # 检查是否是文件对话框（放宽条件）
                                    if w_class == "#32770":
                                        # 如果名称包含"打开"或者是空的，都可能是文件对话框
                                        if "打开" in w_name or w_name == "" or "Open" in w_name:
                                            # 进一步验证：检查是否有文件名输入框
                                            try:
                                                file_edit = w.EditControl(AutomationId="1148")
                                                if file_edit.Exists(maxSearchSeconds=0.3):
                                                    dialog = w
                                                    print(f"  ✓ 找到文件选择对话框: Name='{w_name}' (ClassName: {w_class})，尝试次数: {attempt + 1}")
                                                    break
                                            except:
                                                # 即使找不到输入框，也使用这个窗口（如果类名匹配）
                                                dialog = w
                                                print(f"  ✓ 找到可能的文件选择对话框: Name='{w_name}' (ClassName: {w_class})，尝试次数: {attempt + 1}")
                                                break
                            except:
                                continue
                        
                        if dialog:
                            break
                    except Exception as e:
                        if attempt == 0:
                            print(f"  ⚠️ 遍历窗口时出错: {e}")
                    time.sleep(0.5)
            
            # 方法3: 如果还是找不到，尝试查找所有模态对话框
            if not dialog:
                print("  🔍 方法2未找到，尝试查找所有模态对话框...")
                for attempt in range(10):  # 最多等待5秒
                    try:
                        root = auto.GetRootControl()
                        all_windows = root.GetChildren()
                        for w in all_windows:
                            try:
                                if w.ControlType == auto.ControlType.WindowControl:
                                    # 检查是否是模态对话框
                                    try:
                                        window_pattern = w.GetWindowPattern()
                                        if window_pattern and window_pattern.IsModal:
                                            w_name = (w.Name or "").strip()
                                            w_class = w.ClassName or ""
                                            # 排除主窗口
                                            if w_name != "批量导入" and w_name:
                                                dialog = w
                                                print(f"  ✓ 找到模态对话框: {w_name} (ClassName: {w_class})，尝试次数: {attempt + 1}")
                                                break
                                    except:
                                        pass
                            except:
                                continue
                        
                        if dialog:
                            break
                    except:
                        pass
                    time.sleep(0.5)
            
            if not dialog:
                print("  ⚠️ 未找到文件选择对话框（已等待15秒）")
                print("  💡 提示：请检查文件选择对话框是否已打开")
                return False
            
            print("  ✓ 找到文件选择对话框")
            
            # 3. 构建模板文件名：站点.wpfTpl
            template_file_name = f"{site}.wpfTpl"
            print(f"  📝 准备输入文件名: {template_file_name}")
            
            # 4. 在文件对话框中输入文件名
            # Win32文件对话框的文件名输入框查找方法
            file_name_edit = None
            try:
                # 方法1: 通过AutomationId查找（Win32对话框的标准ID）
                file_name_edit = dialog.EditControl(AutomationId="1148")
                if not file_name_edit.Exists(maxSearchSeconds=1):
                    file_name_edit = None
            except:
                pass
            
            if not file_name_edit:
                try:
                    # 方法2: 查找所有EditControl，找到文件名输入框
                    # 文件名输入框通常在"文件名(N):"标签附近
                    all_edits = []
                    def find_edits(ctrl, depth=0):
                        if depth > 5:
                            return
                        try:
                            if ctrl.ControlType == auto.ControlType.EditControl:
                                all_edits.append(ctrl)
                        except:
                            pass
                        try:
                            child = ctrl.GetFirstChildControl()
                            while child:
                                find_edits(child, depth + 1)
                                child = child.GetNextSiblingControl()
                        except:
                            pass
                    
                    find_edits(dialog)
                    
                    # 文件名输入框通常是第一个或第二个EditControl
                    if all_edits:
                        file_name_edit = all_edits[0]
                        print(f"  ✓ 找到文件名输入框（通过遍历控件）")
                except Exception as e:
                    print(f"  ⚠️ 查找文件名输入框失败: {e}")
            
            if file_name_edit and file_name_edit.Exists(maxSearchSeconds=1):
                try:
                    # 方法1: 尝试使用ValuePattern直接设置值（最可靠）
                    try:
                        value_pattern = file_name_edit.GetValuePattern()
                        if value_pattern:
                            value_pattern.SetValue("")  # 先清空
                            time.sleep(0.2)
                            value_pattern.SetValue(template_file_name)  # 再设置文件名
                            time.sleep(0.3)
                            print(f"  ✓ 已通过ValuePattern在文件名输入框中输入: {template_file_name}")
                        else:
                            raise Exception("无法获取ValuePattern")
                    except:
                        # 方法2: 如果ValuePattern失败，使用键盘输入
                        file_name_edit.SetFocus()
                        time.sleep(0.3)
                        # 使用Windows API发送Ctrl+A和Delete来清空
                        # Ctrl+A
                        user32.keybd_event(0x11, 0, 0, 0)  # VK_CONTROL down
                        user32.keybd_event(0x41, 0, 0, 0)  # VK_A down
                        time.sleep(0.1)
                        user32.keybd_event(0x41, 0, 2, 0)  # VK_A up
                        user32.keybd_event(0x11, 0, 2, 0)  # VK_CONTROL up
                        time.sleep(0.2)
                        # Delete
                        user32.keybd_event(0x2E, 0, 0, 0)  # VK_DELETE down
                        time.sleep(0.05)
                        user32.keybd_event(0x2E, 0, 2, 0)  # VK_DELETE up
                        time.sleep(0.2)
                        # 输入文件名
                        auto.SendKeys(template_file_name)
                        time.sleep(0.5)
                        print(f"  ✓ 已通过键盘输入在文件名输入框中输入: {template_file_name}")
                except Exception as e:
                    print(f"  ⚠️ 输入文件名失败: {e}")
                    # 如果SetFocus失败，尝试直接使用SendKeys到对话框
                    try:
                        dialog.SetFocus()
                        time.sleep(0.3)
                        # 使用Windows API发送Ctrl+A和Delete
                        user32.keybd_event(0x11, 0, 0, 0)  # VK_CONTROL down
                        user32.keybd_event(0x41, 0, 0, 0)  # VK_A down
                        time.sleep(0.1)
                        user32.keybd_event(0x41, 0, 2, 0)  # VK_A up
                        user32.keybd_event(0x11, 0, 2, 0)  # VK_CONTROL up
                        time.sleep(0.2)
                        user32.keybd_event(0x2E, 0, 0, 0)  # VK_DELETE down
                        time.sleep(0.05)
                        user32.keybd_event(0x2E, 0, 2, 0)  # VK_DELETE up
                        time.sleep(0.2)
                        auto.SendKeys(template_file_name)
                        time.sleep(0.5)
                        print(f"  ✓ 已通过对话框焦点输入文件名")
                    except:
                        pass
            else:
                print("  ⚠️ 未找到文件名输入框，尝试直接发送文件名到对话框...")
                try:
                    dialog.SetFocus()
                    time.sleep(0.3)
                    # 使用Windows API发送Ctrl+A和Delete
                    user32.keybd_event(0x11, 0, 0, 0)  # VK_CONTROL down
                    user32.keybd_event(0x41, 0, 0, 0)  # VK_A down
                    time.sleep(0.1)
                    user32.keybd_event(0x41, 0, 2, 0)  # VK_A up
                    user32.keybd_event(0x11, 0, 2, 0)  # VK_CONTROL up
                    time.sleep(0.2)
                    user32.keybd_event(0x2E, 0, 0, 0)  # VK_DELETE down
                    time.sleep(0.05)
                    user32.keybd_event(0x2E, 0, 2, 0)  # VK_DELETE up
                    time.sleep(0.2)
                    auto.SendKeys(template_file_name)
                    time.sleep(0.5)
                    print(f"  ✓ 已直接输入文件名到对话框: {template_file_name}")
                except Exception as e:
                    print(f"  ⚠️ 直接输入文件名失败: {e}")
            
            # 5. 点击"打开"按钮
            print("  🔍 开始查找'打开'按钮...")
            open_button = None
            
            # 先等待一下，确保对话框完全加载
            time.sleep(0.5)
            
            # 方法1: 通过AutomationId查找（Win32对话框的标准ID，最可靠）
            try:
                print("  🔍 方法1: 通过AutomationId='1'查找...")
                open_button = dialog.ButtonControl(AutomationId="1")
                if open_button and open_button.Exists(maxSearchSeconds=1):
                    btn_name = (open_button.Name or "").strip()
                    print(f"  ✓ 找到'打开'按钮（通过AutomationId='1'），名称: '{btn_name}'")
                else:
                    open_button = None
            except Exception as e:
                print(f"  ⚠️ 通过AutomationId查找失败: {e}")
                open_button = None
            
            # 方法2: 通过Name查找（Win32对话框）
            if not open_button:
                try:
                    print("  🔍 方法2: 通过Name查找...")
                    # 按照用户提供的信息，按钮名称是"打开(O)"
                    open_button = dialog.ButtonControl(Name="打开(O)")
                    if not open_button.Exists(maxSearchSeconds=1):
                        open_button = dialog.ButtonControl(Name="打开(&O)")
                        if not open_button.Exists(maxSearchSeconds=1):
                            open_button = dialog.ButtonControl(Name="打开")
                            if not open_button.Exists(maxSearchSeconds=1):
                                open_button = dialog.ButtonControl(Name="Open")
                    if open_button and open_button.Exists(maxSearchSeconds=0.5):
                        btn_name = (open_button.Name or "").strip()
                        print(f"  ✓ 找到'打开'按钮（通过Name），名称: '{btn_name}'")
                    else:
                        open_button = None
                except Exception as e:
                    print(f"  ⚠️ 通过Name查找失败: {e}")
                    open_button = None
            
            # 方法3: 遍历所有按钮，找到"打开"按钮
            if not open_button:
                try:
                    print("  🔍 方法3: 遍历所有按钮...")
                    def find_open_button(ctrl, depth=0):
                        if depth > 10:  # 增加深度限制
                            return None
                        try:
                            if ctrl.ControlType == auto.ControlType.ButtonControl:
                                btn_name = (ctrl.Name or "").strip()
                                if btn_name:
                                    print(f"    发现按钮: '{btn_name}'")
                                if "打开" in btn_name or "Open" in btn_name:
                                    return ctrl
                        except:
                            pass
                        try:
                            child = ctrl.GetFirstChildControl()
                            while child:
                                result = find_open_button(child, depth + 1)
                                if result:
                                    return result
                                child = child.GetNextSiblingControl()
                        except:
                            pass
                        return None
            
                    open_button = find_open_button(dialog)
                    if open_button:
                        btn_name = (open_button.Name or "").strip()
                        print(f"  ✓ 找到'打开'按钮（通过遍历控件），名称: '{btn_name}'")
                except Exception as e:
                    print(f"  ⚠️ 遍历控件查找失败: {e}")
            
            # 方法4: 通过GetChildren查找所有按钮
            if not open_button:
                try:
                    print("  🔍 方法4: 通过GetChildren查找所有按钮...")
                    all_buttons = []
                    def collect_buttons(ctrl, depth=0):
                        if depth > 10:
                            return
                        try:
                            if ctrl.ControlType == auto.ControlType.ButtonControl:
                                all_buttons.append(ctrl)
                        except:
                            pass
                        try:
                            children = ctrl.GetChildren()
                            for child in children:
                                collect_buttons(child, depth + 1)
                        except:
                            pass
                    
                    collect_buttons(dialog)
                    print(f"    找到 {len(all_buttons)} 个按钮")
                    for btn in all_buttons:
                        try:
                            btn_name = (btn.Name or "").strip()
                            btn_id = btn.AutomationId or ""
                            print(f"    按钮: Name='{btn_name}', AutomationId='{btn_id}'")
                            if "打开" in btn_name or "Open" in btn_name or btn_id == "1":
                                open_button = btn
                                print(f"  ✓ 找到'打开'按钮（通过GetChildren），名称: '{btn_name}'")
                                break
                        except:
                            pass
                except Exception as e:
                    print(f"  ⚠️ 通过GetChildren查找失败: {e}")
            
            button_clicked = False  # 标记是否成功点击了按钮
            
            if open_button and open_button.Exists(maxSearchSeconds=1):
                try:
                    open_button.Click()
                    print("  ✓ 已点击'打开'按钮")
                    button_clicked = True
                    human_delay(1.0, 1.5)
                except Exception as e:
                    print(f"  ⚠️ Click失败: {e}，尝试其他方法...")
                    # 如果Click失败，尝试使用InvokePattern
                    try:
                        invoke_pattern = open_button.GetInvokePattern()
                        if invoke_pattern:
                            invoke_pattern.Invoke()
                            print("  ✓ 已通过InvokePattern点击'打开'按钮")
                            button_clicked = True
                            human_delay(1.0, 1.5)
                    except Exception as e2:
                        print(f"  ⚠️ InvokePattern失败: {e2}")
                        # 如果InvokePattern也失败，尝试坐标点击
                        if not button_clicked:
                            try:
                                rect = open_button.BoundingRectangle
                                if rect.width() > 0 and rect.height() > 0:
                                    cx = rect.left + rect.width() // 2
                                    cy = rect.top + rect.height() // 2
                                    if mouse_click(cx, cy):
                                        print("  ✓ 已通过坐标点击'打开'按钮")
                                        button_clicked = True
                                        human_delay(1.0, 1.5)
                            except Exception as e3:
                                print(f"  ⚠️ 坐标点击失败: {e3}")
            
            # 如果已经成功点击了按钮，等待1秒后按Enter键处理弹窗
            if button_clicked:
                print("  ⏳ 等待1秒后按Enter键处理弹窗...")
                time.sleep(1.0)  # 等待1秒
                
                try:
                    # 直接按Enter键
                    auto.SendKeys("{ENTER}")
                    print("  ✓ 已按Enter键处理弹窗")
                    human_delay(0.5, 0.8)
                except Exception as e:
                    print(f"  ⚠️ 按Enter键失败: {e}")
                
                return True
            
            # 如果找不到打开按钮或点击失败，尝试按Enter键
            if not open_button or not open_button.Exists(maxSearchSeconds=0.5):
                print("  ⚠️ 未找到'打开'按钮，尝试按Enter键...")
                try:
                    dialog.SetFocus()
                    time.sleep(0.3)
                    auto.SendKeys("{ENTER}")
                    print("  ✓ 已按Enter键确认")
                    human_delay(1.0, 1.5)
                except Exception as e:
                    print(f"  ❌ 按Enter键失败: {e}")
                    return False
            
            return True
                
        except Exception as e:
            print(f"  ❌ 导入模板操作失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def click_size_image_button(self, size_image_file: str) -> bool:
        """
        点击"尺寸图：未设置"按钮，并在文件选择对话框中选择对应的文件
        :param size_image_file: 尺寸图文件名（从表格"尺寸图"列读取）
        :return: 是否成功完成尺寸图设置操作
        """
        if not self.window:
            print("❌ 窗口未初始化，无法点击尺寸图按钮")
            return False
        
        print(f"\n🖼️ 开始设置尺寸图...")
        print(f"   文件: {size_image_file}")
        
        try:
            # 1. 查找并点击"尺寸图片：未设置"按钮
            size_button = self.window.ButtonControl(AutomationId="PictureSizeSkinButton")
            if not size_button.Exists(maxSearchSeconds=2):
                print("  ⚠️ 未找到'尺寸图片：未设置'按钮")
                return False
            
            print("  ✓ 找到'尺寸图片：未设置'按钮")
            
            # 点击按钮
            try:
                size_button.Click()
                print("  ✓ 已点击尺寸图按钮")
            except:
                try:
                    invoke_pattern = size_button.GetInvokePattern()
                    if invoke_pattern:
                        invoke_pattern.Invoke()
                        print("  ✓ 已通过InvokePattern点击尺寸图按钮")
                except:
                    print("  ❌ 点击尺寸图按钮失败")
                    return False
            
            human_delay(1.5, 2.0)  # 等待对话框出现
            
            # 2. 复用共享的文件对话框逻辑（与 PDF 相同）
            if handle_file_select_dialog(size_image_file, wait_after_click=0):
                print("  ✓ 尺寸图设置完成")
                return True
            return False
                
        except Exception as e:
            print(f"  ❌ 设置尺寸图操作失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def configure_upload_settings(self, warehouse_fee: str = "10") -> bool:
        """
        配置产品上传设置区域的按钮和输入框
        :param warehouse_fee: 仓储费，从价格.json获取，默认为"10"
        """
        if not self.window:
            print("❌ 窗口未初始化，无法配置上传设置")
            return False

        print(f"\n⚙️ 开始配置上传设置...")

        # 需要勾选的控件
        controls_to_check = [
            ("NonPictureFirst", "是否规格非图片优先"),
            ("DianPuskinRadioButton", "店铺"),
            ("IsRemoveChinese", "是否删除中文"),
        ]
        # 需要取消勾选的控件
        controls_to_uncheck = [
            ("DeleteFirstImgskinCheckBox", "删除首图"),
            ("ImgSJskinCheckBox", "图片顺序随机"),
        ]
        # 需要填写的输入框
        fields_to_fill = [
            ("HotSearchCountskinTextBox", "热搜词次数", "3"),
            ("Long", "长", "15"),
            ("Width", "宽", "10"),
            ("High", "高", "5"),
            ("FavoriteskinTextBox", "仓储费", warehouse_fee),
        ]

        success = True

        # 勾选控件
        for auto_id, label in controls_to_check:
            try:
                ctrl = self.window.Control(AutomationId=auto_id)
                if ctrl.Exists(maxSearchSeconds=2):
                    try:
                        toggle = ctrl.GetTogglePattern()
                        if toggle and toggle.ToggleState == 1:
                            print(f"  ✓ 「{label}」已勾选，跳过")
                            continue
                    except:
                        pass
                    ctrl.Click()
                    print(f"  ✓ 已勾选「{label}」")
                    time.sleep(0.3)
                else:
                    print(f"  ⚠️ 未找到「{label}」(AutomationId={auto_id})")
                    success = False
            except Exception as e:
                print(f"  ⚠️ 勾选「{label}」失败: {e}")
                success = False

        # 取消勾选控件
        for auto_id, label in controls_to_uncheck:
            try:
                ctrl = self.window.Control(AutomationId=auto_id)
                if ctrl.Exists(maxSearchSeconds=2):
                    try:
                        toggle = ctrl.GetTogglePattern()
                        if toggle and toggle.ToggleState == 0:
                            print(f"  ✓ 「{label}」已取消勾选，跳过")
                            continue
                    except:
                        pass
                    ctrl.Click()
                    print(f"  ✓ 已取消勾选「{label}」")
                    time.sleep(0.3)
                else:
                    print(f"  ⚠️ 未找到「{label}」(AutomationId={auto_id})")
                    success = False
            except Exception as e:
                print(f"  ⚠️ 取消勾选「{label}」失败: {e}")
                success = False

        # 填写输入框
        for auto_id, label, value in fields_to_fill:
            try:
                edit = self.window.EditControl(AutomationId=auto_id)
                if edit.Exists(maxSearchSeconds=2):
                    try:
                        vp = edit.GetValuePattern()
                        if vp:
                            vp.SetValue(value)
                            print(f"  ✓ 已设置「{label}」= {value}")
                            time.sleep(0.2)
                            continue
                    except:
                        pass
                    edit.Click()
                    time.sleep(0.2)
                    auto.SendKeys("{Ctrl}a")
                    time.sleep(0.1)
                    auto.SendKeys(value)
                    print(f"  ✓ 已通过键盘输入「{label}」= {value}")
                    time.sleep(0.2)
                else:
                    print(f"  ⚠️ 未找到「{label}」(AutomationId={auto_id})")
                    success = False
            except Exception as e:
                print(f"  ⚠️ 设置「{label}」失败: {e}")
                success = False

        if success:
            print("  ✅ 上传设置配置完成")
        else:
            print("  ⚠️ 部分上传设置配置失败，但继续执行")
        return success

    def select_template(self) -> bool:
        """
        点击「选择模板」，在「模板编辑」对话框的下拉框中选择「TT 模板」后确认。
        使用 uiautomation + 有界深度遍历：避免嵌套跨进程 UIA 下 pywinauto 全树 descendants 卡死。
        :return: 是否成功完成模板选择操作
        """
        print(f"\n📝 开始选择模板...")
        
        # 检查窗口是否已初始化
        if not self.window:
            print("❌ 窗口未初始化，无法选择模板")
            return False
        
        # === 步骤1: 使用已有窗口，设置焦点 ===
        print("🔍 设置窗口焦点...")
        try:
            # 使用已有的window引用，设置焦点
            self.window.SetFocus()
            time.sleep(0.3)
        except Exception as e:
            print(f"❌ 设置窗口焦点失败: {e}")
            return False

        # === 步骤2: 勾选「原有描述+自定义模板」「英语」「是否店铺翻译」 ===
        buttons_to_select = [
            ("OLDZDYskinRadioButton3", "原有描述+自定义模板"),
            ("EnskinRadioButton", "英语"),
            ("IsShopFy", "是否店铺翻译"),
        ]
        for auto_id, label in buttons_to_select:
            try:
                ctrl = self.window.Control(AutomationId=auto_id)
                if ctrl.Exists(maxSearchSeconds=2):
                    ctrl.Click()
                    print(f"  ✓ 已选中「{label}」(AutomationId={auto_id})")
                    time.sleep(0.3)
                else:
                    print(f"  ⚠️ 未找到「{label}」(AutomationId={auto_id})，继续执行")
            except Exception as e:
                print(f"  ⚠️ 选中「{label}」失败: {e}，继续执行")

        # === 步骤3: 点击「选择模板」按钮 ===
        try:
            # 使用AutomationId查找并点击按钮（与其他操作保持一致）
            select_btn = self.window.ButtonControl(AutomationId="SelectTemplateButton")
            if not select_btn.Exists(maxSearchSeconds=2):
                # 尝试通过名称查找
                select_btn = self.window.ButtonControl(Name="选择模板")
                if not select_btn.Exists(maxSearchSeconds=2):
                    print("❌ 未找到「选择模板」按钮")
                    return False
            
            # 点击按钮
            try:
                select_btn.Click()
                print("✅ 已点击「选择模板」按钮")
            except:
                # 如果Click失败，尝试使用InvokePattern
                try:
                    invoke_pattern = select_btn.GetInvokePattern()
                    if invoke_pattern:
                        invoke_pattern.Invoke()
                        print("✅ 已通过InvokePattern点击「选择模板」按钮")
                except:
                    print("❌ 点击「选择模板」按钮失败")
                    return False
        except Exception as e:
            print(f"❌ 找不到「选择模板」按钮: {e}")
            return False

        print("  ⏳ 等待「模板编辑」对话框…", flush=True)
        time.sleep(1.2)  # 等待模板窗口弹出

        # === 步骤4: 定位「模板编辑」（轮询超时，避免无限等待）===
        template_dlg = _wait_template_edit_window(timeout_sec=45.0)
        if not template_dlg:
            print("❌ 超时未找到「模板编辑」窗口（请确认弹窗未被遮挡、标题仍为「模板编辑」）")
            return False
        try:
            template_dlg.SetFocus()
        except Exception:
            pass
        print("✅ 已打开「模板编辑」窗口", flush=True)

        # === 步骤5: 下拉框选择「TT 模板」===
        try:
            template_combo = _first_combo_under(template_dlg)
            if not template_combo:
                print("❌ 未找到模板下拉框 (ComboBox)")
                return False
            print("✅ 找到模板下拉框", flush=True)

            template_combo.Click()
            time.sleep(0.8)

            target_item = _first_list_item_named(template_dlg, "TT 模板")
            if not target_item:
                print("  ℹ️ 在对话框内未找到列表项，尝试浅扫顶层窗口（下拉浮层）…", flush=True)
                target_item = _first_list_item_named_under_top_windows("TT 模板")

            if not target_item:
                print("❌ 未找到「TT模板」列表项")
                return False

            print("✅ 找到「TT模板」，正在选择...", flush=True)
            target_item.Click()
            time.sleep(0.5)

        except Exception as e:
            print(f"❌ 操作模板下拉框失败: {e}")
            return False

        # === 步骤6: 点击「确认」===
        try:
            confirm_btn = template_dlg.ButtonControl(Name="确认")
            if not confirm_btn.Exists(maxSearchSeconds=4):
                print("❌ 找不到「确认」按钮")
                return False
            try:
                confirm_btn.Click()
            except Exception:
                invoke_pattern = confirm_btn.GetInvokePattern()
                if invoke_pattern:
                    invoke_pattern.Invoke()
                else:
                    raise
            print("✅ 已点击「确认」按钮", flush=True)
        except Exception as e:
            print(f"❌ 点击「确认」按钮失败: {e}")
            return False

        time.sleep(0.6)

        # === 步骤7: 确保焦点回到主窗口（可选）===
        try:
            # 使用已有的window引用设置焦点
            self.window.SetFocus()
            print("✅ 已返回主窗口")
        except:
            pass

        # 滚动到最底部，确保后续控件可见
        try:
            self.window.SetFocus()
            time.sleep(0.3)
            window_rect = self.window.BoundingRectangle
            scroll_x = window_rect.left + window_rect.width() // 2
            scroll_y = window_rect.top + window_rect.height() * 2 // 3
            user32.SetCursorPos(int(scroll_x), int(scroll_y))
            time.sleep(0.2)
            for _ in range(30):
                user32.mouse_event(0x0800, 0, 0, -120, 0)
                time.sleep(0.1)
            print("  ✓ 已滚动到最底部")
            time.sleep(0.5)
        except Exception as e:
            print(f"  ⚠️ 滚动到底部失败: {e}")

        print("🎉 模板选择流程成功完成！")
        return True
    
    def _get_dropdown_item_text(self, list_item) -> str:
        """
        获取下拉框列表项的显示文本
        :param list_item: 列表项控件
        :return: 显示文本
        """
        try:
            # 方法1: 从子控件中获取文本
            child = list_item.GetFirstChildControl()
            while child:
                if child.ControlType == auto.ControlType.TextControl:
                    text = (child.Name or "").strip()
                    if text and not text.startswith("SHOPEE.") and len(text) <= 100:
                        return text
                child = child.GetNextSiblingControl()
        except:
            pass
        
        # 方法2: 尝试从Name属性获取
        try:
            text = (list_item.Name or "").strip()
            if text and not text.startswith("SHOPEE.") and len(text) <= 100:
                return text
        except:
            pass
        
        return ""
