"""
UI自动化辅助工具函数
"""

import time
import random
import uiautomation as auto
import ctypes


def human_delay(min_s=0.8, max_s=2.5):
    """模拟人类操作的延迟"""
    time.sleep(random.uniform(min_s, max_s))


def mouse_click(x, y, button="left"):
    """鼠标点击
    
    Args:
        x: X坐标
        y: Y坐标
        button: 点击的鼠标按钮 ("left", "right", "middle")
    
    Returns:
        bool: 是否成功点击
    """
    if x < 50 or y < 50 or x > 3000 or y > 2000:
        return False
    
    try:
        user32 = ctypes.windll.user32
        
        # 设置鼠标位置
        user32.SetCursorPos(int(x), int(y))
        time.sleep(0.1)  # 增加等待时间，确保鼠标移动到位
        
        # 根据按钮类型选择事件
        if button == "left":
            down_event = 0x0002  # MOUSEEVENTF_LEFTDOWN
            up_event = 0x0004    # MOUSEEVENTF_LEFTUP
        elif button == "right":
            down_event = 0x0008  # MOUSEEVENTF_RIGHTDOWN
            up_event = 0x0010    # MOUSEEVENTF_RIGHTUP
        elif button == "middle":
            down_event = 0x0020  # MOUSEEVENTF_MIDDLEDOWN
            up_event = 0x0040    # MOUSEEVENTF_MIDDLEUP
        else:
            down_event = 0x0002  # 默认左键
            up_event = 0x0004
        
        # 发送鼠标按下事件
        user32.mouse_event(down_event, 0, 0, 0, 0)
        time.sleep(0.05)  # 增加按下和释放之间的间隔
        
        # 发送鼠标释放事件
        user32.mouse_event(up_event, 0, 0, 0, 0)
        time.sleep(0.05)  # 增加释放后的等待时间
        
        return True
    except Exception as e:
        print(f"鼠标点击出错: {e}")
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


def get_all_descendants(ctrl, depth=0, max_depth=8):
    """获取所有后代控件"""
    if depth > max_depth:
        return []
    descendants = []
    try:
        child = ctrl.GetFirstChildControl()
        while child:
            descendants.append(child)
            descendants.extend(get_all_descendants(child, depth + 1, max_depth))
            child = child.GetNextSiblingControl()
    except:
        pass
    return descendants


def take_snapshot(window):
    """拍摄窗口快照"""
    fingerprints = set()
    all_controls = get_all_descendants(window)
    for ctrl in all_controls:
        try:
            rect = ctrl.BoundingRectangle
            fp = (
                ctrl.ControlType,
                (ctrl.Name or "").strip(),
                round(rect.left), round(rect.top),
                round(rect.width()), round(rect.height())
            )
            fingerprints.add(fp)
        except:
            continue
    return fingerprints




