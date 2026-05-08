"""
类目操作处理模块
包含：点击添加类目按钮等功能
"""

import uiautomation as auto
from .ui_utils import human_delay, safe_click_control


class CategoryHandler:
    """类目操作处理器"""
    
    def __init__(self, window=None, window_handler=None):
        """
        初始化类目处理器
        :param window: 主窗口控件
        :param window_handler: 窗口处理器（用于调用scroll_window_to_top）
        """
        self.window = window
        self.window_handler = window_handler
    
    def click_category_modify_button(self) -> bool:
        """
        点击"添加类目"按钮，打开类目修改窗口
        先向上滚动窗口，确保按钮可见，然后再点击
        :return: 是否成功点击按钮
        """
        if not self.window:
            print("❌ 窗口未初始化，无法点击添加类目按钮")
            return False
        
        print("\n🔘 查找并点击'添加类目'按钮...")
        
        try:
            # 先向上滚动窗口到最顶部，确保按钮可见
            if self.window_handler:
                self.window_handler.scroll_window_to_top()
            human_delay(0.5, 0.8)
            
            # 获取窗口位置，用于排除标题栏按钮
            window_rect = self.window.BoundingRectangle
            title_bar_height = 30  # 标题栏大约高度，用于排除窗口控制按钮
            
            # 排除的按钮名称（窗口控制按钮）
            excluded_names = ["最小化", "最大化", "关闭", "Minimize", "Maximize", "Close", 
                            "还原", "Restore", "还原窗口", "最大化窗口", "最小化窗口"]
            
            # 方法1: 优先使用AutomationId查找（最可靠）
            button = None
            try:
                button = self.window.ButtonControl(AutomationId="TypeskinButto")
                if button.Exists(maxSearchSeconds=2):
                    print(f"  ✓ 通过AutomationId找到按钮: TypeskinButto")
                else:
                    button = None
            except:
                button = None
            
            # 方法2: 精确查找"添加类目"按钮，排除窗口控制按钮
            if not button:
                def search_category_button(ctrl, depth=0):
                    if depth > 10:
                        return None
                    try:
                        if ctrl.ControlType == auto.ControlType.ButtonControl:
                            name = (ctrl.Name or "").strip()
                            
                            # 排除窗口控制按钮
                            if any(excluded in name for excluded in excluded_names):
                                return None
                            
                            # 精确匹配"添加类目"
                            if name == "添加类目":
                                # 检查按钮位置，确保不在标题栏
                                try:
                                    button_rect = ctrl.BoundingRectangle
                                    # 如果按钮在标题栏区域（窗口顶部30像素内），跳过
                                    if button_rect.top < window_rect.top + title_bar_height:
                                        print(f"  ⚠️ 跳过标题栏按钮: {name}")
                                        return None
                                except:
                                    pass
                                return ctrl
                    except:
                        pass
                    
                    try:
                        child = ctrl.GetFirstChildControl()
                        while child:
                            result = search_category_button(child, depth + 1)
                            if result:
                                return result
                            child = child.GetNextSiblingControl()
                    except:
                        pass
                    return None
                
                button = search_category_button(self.window)
            
            # 方法3: 如果精确匹配失败，尝试查找包含"添加"和"类目"的按钮（但排除窗口控制按钮）
            if not button:
                def search_buttons_with_exclude(ctrl, depth=0):
                    if depth > 10:
                        return None
                    try:
                        if ctrl.ControlType == auto.ControlType.ButtonControl:
                            name = (ctrl.Name or "").strip()
                            
                            # 排除窗口控制按钮
                            if any(excluded in name for excluded in excluded_names):
                                return None
                            
                            # 必须同时包含"添加"和"类目"
                            if "添加" in name and "类目" in name:
                                # 检查按钮位置，确保不在标题栏
                                try:
                                    button_rect = ctrl.BoundingRectangle
                                    if button_rect.top < window_rect.top + title_bar_height:
                                        print(f"  ⚠️ 跳过标题栏按钮: {name}")
                                        return None
                                except:
                                    pass
                                return ctrl
                    except:
                        pass
                    
                    try:
                        child = ctrl.GetFirstChildControl()
                        while child:
                            result = search_buttons_with_exclude(child, depth + 1)
                            if result:
                                return result
                            child = child.GetNextSiblingControl()
                    except:
                        pass
                    return None
                
                button = search_buttons_with_exclude(self.window)
            
            if button:
                button_name = button.Name or ""
                print(f"  ✓ 找到按钮: {button_name}")
                
                # 再次检查按钮是否在可见区域
                try:
                    button_rect = button.BoundingRectangle
                    
                    # 如果按钮不在可见区域，继续向上滚动
                    if button_rect.top < window_rect.top or button_rect.bottom > window_rect.bottom:
                        print(f"  ↕️ 按钮不在可见区域，再次滚动到顶部...")
                        if self.window_handler:
                            self.window_handler.scroll_window_to_top()
                        human_delay(0.5, 0.8)
                    
                    # 打印按钮位置信息（用于调试）
                    print(f"    按钮位置: ({button_rect.left}, {button_rect.top}), 大小: {button_rect.width()}x{button_rect.height()}")
                except:
                    pass
                
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
                # 打印所有按钮名称用于调试
                try:
                    print("  🔍 窗口中的所有按钮:")
                    def print_all_buttons(ctrl, depth=0):
                        if depth > 5:
                            return
                        try:
                            if ctrl.ControlType == auto.ControlType.ButtonControl:
                                name = (ctrl.Name or "").strip()
                                if name:
                                    try:
                                        rect = ctrl.BoundingRectangle
                                        print(f"    - {name} (位置: {rect.top})")
                                    except:
                                        print(f"    - {name}")
                        except:
                            pass
                        try:
                            child = ctrl.GetFirstChildControl()
                            while child:
                                print_all_buttons(child, depth + 1)
                                child = child.GetNextSiblingControl()
                        except:
                            pass
                    print_all_buttons(self.window)
                except:
                    pass
                return False
                
        except Exception as e:
            print(f"  ❌ 查找或点击按钮时出错: {e}")
            return False
