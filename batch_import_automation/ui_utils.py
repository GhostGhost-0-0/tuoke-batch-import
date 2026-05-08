"""
UI自动化辅助工具函数
"""

import time
import random
import ctypes
import uiautomation as auto

# Windows API 用于鼠标事件
user32 = ctypes.windll.user32

# 全局暂停事件（从外部设置）
_global_pause_event = None


def set_global_pause_event(pause_event):
    """设置全局暂停事件"""
    global _global_pause_event
    _global_pause_event = pause_event


def human_delay(min_s=0.8, max_s=2.5):
    """
    模拟人类操作的延迟，支持暂停检查
    :param min_s: 最小延迟时间（秒）
    :param max_s: 最大延迟时间（秒）
    """
    delay_time = random.uniform(min_s, max_s)

    # 如果有暂停事件，分小段等待并检查暂停状态
    if _global_pause_event:
        segment_time = 0.1  # 每段等待0.1秒
        elapsed = 0

        while elapsed < delay_time:
            # 检查是否需要暂停
            if _global_pause_event.is_set():
                print("⏸️ 延迟中，等待恢复...")
                _global_pause_event.wait()  # 阻塞等待直到事件被清除
                print("▶️ 已恢复延迟")

            # 等待一小段时间
            remaining = min(segment_time, delay_time - elapsed)
            time.sleep(remaining)
            elapsed += remaining
    else:
        # 如果没有暂停事件，直接等待
        time.sleep(delay_time)


def mouse_click(x, y):
    """
    通过坐标点击鼠标
    :param x: X坐标
    :param y: Y坐标
    :return: 是否成功点击
    """
    try:
        # 验证坐标范围（防止无效坐标）
        if x < 50 or y < 50 or x > 3000 or y > 2000:
            return False
        
        # 移动鼠标到指定位置
        user32.SetCursorPos(int(x), int(y))
        time.sleep(0.05)
        
        # 按下鼠标左键
        user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFT DOWN
        time.sleep(0.03)
        
        # 释放鼠标左键
        user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFT UP
        return True
    except Exception as e:
        print(f"鼠标点击失败: {e}")
        return False


def safe_click_control(control, timeout=2):
    """
    安全地点击控件
    :param control: 控件对象
    :param timeout: 超时时间（秒）
    :return: 是否成功点击
    """
    try:
        if control.Exists(maxSearchSeconds=timeout):
            control.Click()
            return True
    except Exception as e:
        print(f"点击控件失败: {e}")
    return False


def set_control_value(control, value):
    """
    设置控件值
    :param control: 控件对象
    :param value: 要设置的值
    :return: 是否成功设置
    """
    try:
        # 方法1: 尝试使用 ValuePattern 直接设置值（最可靠）
        try:
            value_pattern = control.GetValuePattern()
            if value_pattern:
                # 先设置焦点
                control.SetFocus()
                time.sleep(0.15)
                # 直接设置值（不需要先清空，SetValue会覆盖）
                value_str = str(value)
                value_pattern.SetValue(value_str)
                time.sleep(0.2)
                
                # 验证设置是否成功
                try:
                    actual_value = value_pattern.Value or ""
                    if actual_value == value_str:
                        # 设置成功后，移除焦点（按Tab键或点击其他地方）
                        try:
                            auto.SendKeys("{TAB}")  # 按Tab键移除焦点
                            time.sleep(0.1)
                        except:
                            pass
                        return True
                except:
                    pass
                
                # 即使验证失败，也移除焦点
                try:
                    auto.SendKeys("{TAB}")
                    time.sleep(0.1)
                except:
                    pass
                
                return True
        except Exception as e:
            # 如果ValuePattern失败，继续尝试其他方法
            pass
        
        # 方法2: 如果 ValuePattern 不可用，使用键盘输入
        try:
            control.SetFocus()
            time.sleep(0.15)
            
            # 方法2a: 使用Ctrl+A全选（使用SendInput方式）
            import ctypes
            VK_CONTROL = 0x11
            VK_A = 0x41
            VK_DELETE = 0x2E
            
            # 按下Ctrl+A
            ctypes.windll.user32.keybd_event(VK_CONTROL, 0, 0, 0)
            time.sleep(0.05)
            ctypes.windll.user32.keybd_event(VK_A, 0, 0, 0)
            time.sleep(0.05)
            ctypes.windll.user32.keybd_event(VK_A, 0, 2, 0)  # 释放A
            time.sleep(0.05)
            ctypes.windll.user32.keybd_event(VK_CONTROL, 0, 2, 0)  # 释放Ctrl
            time.sleep(0.1)
            
            # 删除选中的内容
            ctypes.windll.user32.keybd_event(VK_DELETE, 0, 0, 0)
            time.sleep(0.05)
            ctypes.windll.user32.keybd_event(VK_DELETE, 0, 2, 0)
            time.sleep(0.1)
            
            # 输入新值（确保是字符串）
            value_str = str(value)
            auto.SendKeys(value_str)
            time.sleep(0.2)
            
            # 设置值后，移除焦点（按Tab键）
            try:
                auto.SendKeys("{TAB}")
                time.sleep(0.1)
            except:
                pass
            
            return True
        except Exception as e:
            # 如果键盘输入失败，尝试最后的方法
            print(f"键盘输入方法失败: {e}")
            pass
        
        # 方法3: 最后的备选方案 - 使用SendKeys但更小心
        try:
            control.SetFocus()
            time.sleep(0.15)
            
            # 使用SendKeys，但先点击控件确保焦点
            control.Click()
            time.sleep(0.1)
            
            # 使用End键移动到末尾，然后Shift+Home全选
            auto.SendKeys("{End}")
            time.sleep(0.05)
            auto.SendKeys("+{Home}")  # Shift+Home
            time.sleep(0.1)
            auto.SendKeys("{DELETE}")
            time.sleep(0.1)
            
            # 输入新值
            value_str = str(value)
            auto.SendKeys(value_str)
            time.sleep(0.2)
            
            # 设置值后，移除焦点（按Tab键）
            try:
                auto.SendKeys("{TAB}")
                time.sleep(0.1)
            except:
                pass
            
            return True
        except Exception as e:
            print(f"设置控件值失败: {e}")
            return False
            
    except Exception as e:
        print(f"设置控件值失败: {e}")
        return False

