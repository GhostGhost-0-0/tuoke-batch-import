"""
窗口操作处理模块
"""

import time
import ctypes
import uiautomation as auto
from .ui_utils import human_delay

# Windows API 用于鼠标事件
user32 = ctypes.windll.user32


class WindowHandler:
    """窗口操作处理类"""
    
    WINDOW_NAME = "批量导入"
    
    def __init__(self, window=None):
        """
        初始化窗口处理器
        :param window: 窗口控件对象
        """
        self.window = window
    
    def find_window(self, timeout: int = 5) -> bool:
        """
        查找并激活目标窗口
        :param timeout: 超时时间（秒）
        :return: 是否成功找到窗口
        """
        self.window = auto.WindowControl(Name=self.WINDOW_NAME, searchDepth=1)
        if not self.window.Exists(maxSearchSeconds=timeout):
            print(f"❌ 未找到 '{self.WINDOW_NAME}' 窗口")
            return False
        
        self.window.SetActive()
        human_delay(1.0, 1.5)
        print(f"✅ 已找到并激活 '{self.WINDOW_NAME}' 窗口")
        return True
    
    def scroll_window_to_top(self, wheel_only: bool = False) -> bool:
        """
        向上滚动窗口到最顶部
        :param wheel_only: 是否仅使用鼠标滚轮滚动（避免快捷键作用到错误焦点控件）
        :return: 是否成功滚动
        """
        try:
            if not self.window:
                print("  ⚠️ 窗口未初始化")
                return False
            
            # 激活窗口并设置焦点
            self.window.SetActive()
            time.sleep(0.3)
            self.window.SetFocus()
            time.sleep(0.3)
            
            # 获取窗口位置
            window_rect = self.window.BoundingRectangle
            if window_rect.width() <= 0 or window_rect.height() <= 0:
                print("  ⚠️ 窗口尺寸无效")
                return False
            
            # 计算滚动位置（靠右区域，尽量命中主滚动区域而非分组列表）
            scroll_x = window_rect.right - 35
            scroll_y = window_rect.top + window_rect.height() // 3  # 窗口上1/3处
            
            # 确保坐标在有效范围内
            scroll_x = max(50, min(scroll_x, 2500))
            scroll_y = max(50, min(scroll_y, 1800))
            
            # 移动鼠标到窗口内
            user32.SetCursorPos(int(scroll_x), int(scroll_y))
            time.sleep(0.2)
            
            print(f"  ↕️ 向上滚动窗口到最顶部...")
            
            if not wheel_only:
                # 方法1: 尝试使用 Ctrl+Home 快捷键直接跳到顶部（最可靠）
                try:
                    print(f"    尝试使用 Ctrl+Home 快捷键...")
                    auto.SendKeys("^{HOME}")  # Ctrl+Home
                    time.sleep(0.5)
                except Exception as e:
                    print(f"    Ctrl+Home 失败: {e}")
                
                # 方法2: 使用 Page Up 键快速滚动
                print(f"    使用 Page Up 键快速滚动...")
                for i in range(15):  # 增加次数
                    try:
                        auto.SendKeys("{PGUP}")
                        time.sleep(0.15)
                    except:
                        pass
            else:
                print("    已启用仅滚轮模式，跳过 Ctrl+Home/PageUp")
            
            # 方法3: 使用滚轮向上滚动（确保到达顶部）
            print(f"    使用滚轮精细滚动...")
            # 确保鼠标在窗口内
            user32.SetCursorPos(int(scroll_x), int(scroll_y))
            time.sleep(0.1)
            
            # 向上滚动多次
            for i in range(30):
                try:
                    # 0x0800 = WM_MOUSEWHEEL, 120 = 向上滚动
                    user32.mouse_event(0x0800, 0, 0, 120, 0)
                    time.sleep(0.15)
                except Exception as e:
                    print(f"    滚轮滚动失败: {e}")
                    break
            
            print(f"  ✅ 滚动完成")
            time.sleep(0.5)  # 等待UI稳定
            return True
            
        except Exception as e:
            print(f"  ⚠️ 滚动窗口失败: {e}")
            import traceback
            traceback.print_exc()
            return False

    def scroll_window_to_bottom(self) -> bool:
        """
        向下滚动窗口到最底部
        :return: 是否成功滚动
        """
        try:
            if not self.window:
                print("  ⚠️ 窗口未初始化")
                return False
            
            # 激活窗口并设置焦点
            self.window.SetActive()
            time.sleep(0.3)
            self.window.SetFocus()
            time.sleep(0.3)
            
            # 获取窗口位置
            window_rect = self.window.BoundingRectangle
            if window_rect.width() <= 0 or window_rect.height() <= 0:
                print("  ⚠️ 窗口尺寸无效")
                return False
            
            # 计算滚动位置（窗口中心偏下）
            scroll_x = window_rect.left + window_rect.width() // 2
            scroll_y = window_rect.top + window_rect.height() * 2 // 3  # 窗口下2/3处
            
            # 确保坐标在有效范围内
            scroll_x = max(50, min(scroll_x, 2500))
            scroll_y = max(50, min(scroll_y, 1800))
            
            # 移动鼠标到窗口内
            user32.SetCursorPos(int(scroll_x), int(scroll_y))
            time.sleep(0.2)
            
            print(f"  ↕️ 向下滚动窗口到最底部...")
            
            # 方法1: 尝试使用 Ctrl+End 快捷键直接跳到底部（最可靠）
            try:
                print(f"    尝试使用 Ctrl+End 快捷键...")
                auto.SendKeys("^{END}")  # Ctrl+End
                time.sleep(0.5)
            except Exception as e:
                print(f"    Ctrl+End 失败: {e}")
            
            # 方法2: 使用 Page Down 键快速滚动
            print(f"    使用 Page Down 键快速滚动...")
            for i in range(15):  # 增加次数
                try:
                    auto.SendKeys("{PGDN}")
                    time.sleep(0.15)
                except:
                    pass
            
            # 方法3: 使用滚轮向下滚动（确保到达底部）
            print(f"    使用滚轮精细滚动...")
            # 确保鼠标在窗口内
            user32.SetCursorPos(int(scroll_x), int(scroll_y))
            time.sleep(0.1)
            
            # 向下滚动多次
            for i in range(30):
                try:
                    # 0x0800 = WM_MOUSEWHEEL, -120 = 向下滚动
                    user32.mouse_event(0x0800, 0, 0, -120, 0)
                    time.sleep(0.15)
                except Exception as e:
                    print(f"    滚轮滚动失败: {e}")
                    break
            
            print(f"  ✅ 滚动完成")
            time.sleep(0.5)  # 等待UI稳定
            return True
            
        except Exception as e:
            print(f"  ⚠️ 滚动窗口失败: {e}")
            import traceback
            traceback.print_exc()
            return False