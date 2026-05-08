"""
对话框处理器模块
用于检测和处理弹窗对话框
"""

import time
import uiautomation as auto
from .ui_utils import mouse_click


def handle_dialog_popup(parent_window, timeout=5):
    """
    检测并处理弹窗对话框，自动点击"确定"按钮
    :param parent_window: 父窗口对象
    :param timeout: 等待弹窗出现的超时时间（秒）
    :return: 是否检测到并处理了弹窗（True=有弹窗并已处理，False=没有弹窗）
    """
    print("\n🔍 检测弹窗对话框...")
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        try:
            # 方法1: 查找所有窗口，找到模态对话框
            dialogs = []
            try:
                # 获取根控件的所有子窗口
                root = auto.GetRootControl()
                all_windows = root.GetChildren()
                for w in all_windows:
                    try:
                        if w.ControlType == auto.ControlType.WindowControl:
                            # 检查是否是模态对话框
                            try:
                                window_pattern = w.GetWindowPattern()
                                if window_pattern and window_pattern.IsModal:
                                    # 检查不是主窗口
                                    w_name = w.Name or ""
                                    if w_name != "类目修改":
                                        dialogs.append(w)
                            except:
                                # 如果无法获取WindowPattern，检查其他特征
                                if (w.FrameworkId == "WPF" and 
                                    w.ClassName == "Window" and
                                    (w.Name or "") != "类目修改"):
                                    dialogs.append(w)
                    except:
                        continue
            except Exception as e:
                pass
            
            # 方法2: 如果没找到，尝试直接查找"确定"按钮（可能在当前活动窗口）
            if not dialogs:
                try:
                    # 查找当前活动的窗口
                    active_window = auto.GetForegroundControl()
                    if active_window and active_window.ControlType == auto.ControlType.WindowControl:
                        active_name = active_window.Name or ""
                        if active_name != "类目修改":
                            # 检查是否有"确定"按钮
                            ok_btn = active_window.ButtonControl(Name="确定")
                            if ok_btn.Exists(maxSearchSeconds=0.3):
                                dialogs.append(active_window)
                except:
                    pass
            
            # 如果找到弹窗
            if dialogs:
                dialog = dialogs[0]
                print(f"✅ 检测到弹窗对话框")
                
                # 查找"确定"按钮（优先级：确定 > 是）
                ok_button = None
                
                # 方法1: 直接通过Name查找"确定"
                try:
                    ok_button = dialog.ButtonControl(Name="确定")
                    if not ok_button.Exists(maxSearchSeconds=0.3):
                        ok_button = None
                except:
                    pass
                
                # 方法2: 如果没找到，遍历所有按钮查找
                if not ok_button:
                    try:
                        def find_button_recursive(ctrl, depth=0):
                            if depth > 5:
                                return None
                            try:
                                if ctrl.ControlType == auto.ControlType.ButtonControl:
                                    btn_name = ctrl.Name or ""
                                    if "确定" in btn_name:
                                        return ctrl
                            except:
                                pass
                            
                            try:
                                child = ctrl.GetFirstChildControl()
                                while child:
                                    result = find_button_recursive(child, depth + 1)
                                    if result:
                                        return result
                                    child = child.GetNextSiblingControl()
                            except:
                                pass
                            return None
                        
                        ok_button = find_button_recursive(dialog)
                    except:
                        pass
                
                # 如果找到"确定"按钮
                if ok_button:
                    print(f"✅ 找到'确定'按钮，正在点击...")
                    try:
                        ok_button.Click()
                        time.sleep(0.5)
                        print(f"✅ 已点击'确定'按钮，弹窗已关闭")
                        return True
                    except:
                        # 如果Click失败，使用坐标点击
                        try:
                            r = ok_button.BoundingRectangle
                            if r.width() > 0 and r.height() > 0:
                                cx = r.left + r.width() // 2
                                cy = r.top + r.height() // 2
                                if mouse_click(cx, cy):
                                    time.sleep(0.5)
                                    print(f"✅ 已通过坐标点击'确定'按钮")
                                    return True
                        except:
                            pass
                else:
                    # 如果没找到"确定"，尝试查找"是"按钮
                    print(f"⚠️ 未找到'确定'按钮，尝试查找'是'按钮...")
                    try:
                        yes_button = dialog.ButtonControl(Name="是")
                        if yes_button.Exists(maxSearchSeconds=0.3):
                            try:
                                yes_button.Click()
                                time.sleep(0.5)
                                print(f"✅ 已点击'是'按钮")
                                return True
                            except:
                                r = yes_button.BoundingRectangle
                                if r.width() > 0 and r.height() > 0:
                                    cx = r.left + r.width() // 2
                                    cy = r.top + r.height() // 2
                                    if mouse_click(cx, cy):
                                        time.sleep(0.5)
                                        print(f"✅ 已通过坐标点击'是'按钮")
                                        return True
                    except:
                        pass
            
            time.sleep(0.2)  # 短暂等待后继续检测
            
        except Exception as e:
            time.sleep(0.2)
            continue
    
    print(f"ℹ️ 未检测到弹窗对话框（{timeout}秒内）")
    return False

