"""
下拉框处理器模块
用于处理下拉框的滚动和选择
"""

import time
import uiautomation as auto
import ctypes
from .ui_utils import human_delay, mouse_click, get_all_descendants


user32 = ctypes.windll.user32


def _normalize_text(text: str) -> str:
    """
    统一处理用于匹配的文本，尽量忽略格式差异
    - 全部转小写
    - 去掉首尾空格
    - 去掉所有空白字符（空格、制表符等）
    - 统一常见符号（中英文括号、引号、连字符等）
    """
    if not text:
        return ""
    t = text.strip().lower()
    # 统一括号
    t = t.replace("（", "(").replace("）", ")")
    # 统一引号
    t = (
        t.replace("“", '"')
        .replace("”", '"')
        .replace("‘", "'")
        .replace("’", "'")
    )
    # 统一连字符（减号/短横线等）
    for ch in ["－", "–", "—"]:
        t = t.replace(ch, "-")
    # 去掉所有空白字符（包括中间的空格）
    t = "".join(c for c in t if not c.isspace())
    return t


def get_display_text(list_item):
    """获取列表项的显示文本"""
    try:
        # 方法1: 先尝试从子控件中获取文本（真正的显示文本通常在子控件中）
        try:
            child = list_item.GetFirstChildControl()
            while child:
                if child.ControlType == auto.ControlType.TextControl:
                    text = (child.Name or "").strip()
                    if text and not text.startswith("SHOPEE.") and len(text) <= 100:
                        return text
                child = child.GetNextSiblingControl()
        except:
            pass
        
        # 方法2: 如果子控件中没有，尝试递归查找所有文本控件
        try:
            def find_text_in_children(ctrl, depth=0):
                if depth > 3:  # 限制深度避免无限递归
                    return ""
                try:
                    child = ctrl.GetFirstChildControl()
                    while child:
                        if child.ControlType == auto.ControlType.TextControl:
                            text = (child.Name or "").strip()
                            if text and not text.startswith("SHOPEE.") and len(text) <= 100:
                                return text
                        # 递归查找
                        result = find_text_in_children(child, depth + 1)
                        if result:
                            return result
                        child = child.GetNextSiblingControl()
                except:
                    pass
                return ""
            
            text = find_text_in_children(list_item)
            if text:
                return text
        except:
            pass
        
        # 方法3: 最后尝试使用Name（但排除SHOPEE开头的）
        text = list_item.Name or ""
        if text.startswith("SHOPEE.") or len(text) > 100:
            return ""
        return text.strip()
    except:
        return ""


def select_dropdown_option_with_scroll(combo_ctrl, target_text, field_name="选项"):
    """
    通过滚动查找并选择下拉框选项
    每滚动一次会出现往下走三个选项，需要持续滚动直到找到目标
    :param combo_ctrl: 下拉框控件
    :param target_text: 目标文本
    :param field_name: 字段名称（用于日志显示）
    :return: 是否成功选择
    """
    print(f"\n🔍 在{field_name}下拉框中查找并点击: '{target_text}'（使用滚动条）")
    
    # === 1. 展开下拉框 ===
    dropdown_button = None
    try:
        child = combo_ctrl.GetFirstChildControl()
        while child:
            if child.ControlType == auto.ControlType.ButtonControl and not child.IsOffscreen:
                dropdown_button = child
                break
            child = child.GetNextSiblingControl()
    except:
        pass
    
    if dropdown_button:
        try:
            dropdown_button.Click()
        except:
            r = dropdown_button.BoundingRectangle
            if r.width() > 0 and r.height() > 0:
                mouse_click(r.left + r.width() // 2, r.top + r.height() // 2)
    else:
        try:
            combo_ctrl.SetFocus()
            human_delay(0.6, 1.2)
            auto.SendKeys("{F4}")
            human_delay(0.7, 1.3)
        except:
            pass
    
    time.sleep(0.8)
    
    # === 2. 获取下拉容器 ===
    scroll_viewer = auto.PaneControl(AutomationId="DropDownScrollViewer")
    if not scroll_viewer.Exists(maxSearchSeconds=2):
        print("❌ 未找到 DropDownScrollViewer")
        return False
    
    sv_rect = scroll_viewer.BoundingRectangle
    if sv_rect.width() <= 0 or sv_rect.height() <= 0:
        print("❌ 下拉容器尺寸无效")
        return False
    
    # 计算滚动位置（确保在下拉框区域内）
    scroll_x = sv_rect.left + sv_rect.width() // 2
    scroll_y = sv_rect.top + min(60, sv_rect.height() // 2)
    # 确保坐标在下拉框区域内
    scroll_x = max(sv_rect.left + 10, min(scroll_x, sv_rect.right - 10))
    scroll_y = max(sv_rect.top + 10, min(scroll_y, sv_rect.bottom - 10))
    user32.SetCursorPos(int(scroll_x), int(scroll_y))
    time.sleep(0.1)
    
    # 保存下拉框区域边界，用于后续验证鼠标位置
    dropdown_left = sv_rect.left
    dropdown_right = sv_rect.right
    dropdown_top = sv_rect.top
    dropdown_bottom = sv_rect.bottom
    
    seen_texts = set()  # 使用文本去重（避免重复处理相同文本的选项）
    seen_runtime_ids = set()  # 使用 RuntimeId 去重（避免重复处理同一个控件）
    # 原始清洗（去掉括号后的说明）
    target_clean_raw = target_text.split('(')[0].strip()
    # 用于匹配的标准化文本
    target_clean = _normalize_text(target_clean_raw)
    max_scrolls = 50
    scrolls_done = 0
    all_scanned_options = []  # 记录所有扫描到的选项
    
    while scrolls_done <= max_scrolls:
        # 先读取当前可见的所有选项文本（在滚动之前）
        items = []
        try:
            child = scroll_viewer.GetFirstChildControl()
            while child:
                if child.ControlType == auto.ControlType.ListItemControl:
                    # 检查是否真正可见（不仅检查IsOffscreen，还要检查是否在可见区域内）
                    try:
                        if not child.IsOffscreen:
                            rect = child.BoundingRectangle
                            # 检查控件是否在下拉框的可见区域内
                            if (rect.width() > 0 and rect.height() > 0 and
                                rect.top >= dropdown_top and rect.bottom <= dropdown_bottom and
                                rect.left >= dropdown_left and rect.right <= dropdown_right):
                                items.append(child)
                                # 限制最多12个可见项
                                if len(items) >= 12:
                                    break
                    except:
                        pass
                child = child.GetNextSiblingControl()
        except:
            items = []
        
        if not items and scrolls_done == 0:
            print("⚠️ 初始无任何选项")
            break
        
        # 读取当前帧的所有选项文本
        current_frame_options = []  # 当前帧的选项
        new_items_in_frame = 0  # 当前帧中的新项数量
        
        print(f"  🔍 第 {scrolls_done + 1} 次读取，找到 {len(items)} 个列表项（限制最多12个可见项）")
        
        for idx, item in enumerate(items):
            try:
                # 先获取文本
                text = get_display_text(item).strip()
                if not text:
                    text = (item.Name or "").strip()
                    if text.startswith("SHOPEE.") or len(text) > 100:
                        text = ""
                
                # 打印每个项的详细信息
                try:
                    rect = item.BoundingRectangle
                    print(f"    项 {idx + 1}: 文本='{text}', 位置=({rect.left}, {rect.top}), 尺寸={rect.width()}x{rect.height()}, IsOffscreen={item.IsOffscreen}")
                except:
                    print(f"    项 {idx + 1}: 文本='{text}', 无法获取位置信息")
                
                if not text:
                    continue
                
                # 使用文本去重（避免重复处理相同文本）
                text_norm = _normalize_text(text)
                if text_norm in seen_texts:
                    print(f"      跳过：文本 '{text}' 已处理过（标准化='{text_norm}'）")
                    continue  # 已处理过的文本，跳过
                
                # 使用 RuntimeId 去重（避免重复处理同一个控件）
                runtime_id = None
                try:
                    runtime_id = tuple(item.GetRuntimeId())
                except:
                    rect = item.BoundingRectangle
                    if rect.width() <= 0 or rect.height() <= 0:
                        continue
                    runtime_id = (rect.left, rect.top, rect.width(), rect.height())
                
                if runtime_id in seen_runtime_ids:
                    print(f"      跳过：控件已处理过")
                    continue  # 已处理过的控件，跳过
                
                # 记录已处理的文本和控件
                seen_texts.add(text_norm)
                seen_runtime_ids.add(runtime_id)
                new_items_in_frame += 1
                
                # 记录选项文本
                current_frame_options.append(text)
                if text not in all_scanned_options:
                    all_scanned_options.append(text)
                
                print(f"      ✓ 新项：'{text}' (标准化: '{text_norm}')")
                
                # 检查是否匹配目标（精确匹配，避免"TT女装"误匹配"TT女装2"等情况）
                # 使用标准化后的文本进行比较，尽量忽略空格、引号、连字符等差异
                is_match = (target_clean == text_norm)
                if is_match:
                    print(f"✅ 找到{field_name} '{text}'，准备选择...")
                    
                    # 注意：由于我们已经在滚动列表中查找，找到的项应该已经是可见的
                    # 如果确实需要滚动，可以尝试使用ScrollItemPattern，但通常不需要
                    
                    # 方法1: 使用 SelectionItemPattern 选择（推荐，不需要坐标和键盘）
                    try:
                        selection_pattern = item.GetSelectionItemPattern()
                        if selection_pattern:
                            print(f"  使用 SelectionItemPattern 选择 '{text}'...")
                            selection_pattern.Select()
                            time.sleep(0.5)
                            print(f"✅ 已通过 SelectionItemPattern 选择 '{text}'")
                            return True
                    except Exception as e:
                        print(f"  SelectionItemPattern 不可用: {e}")
                    
                    # 方法2: 直接调用控件的 Click() 方法
                    try:
                        print(f"  尝试直接点击控件 '{text}'...")
                        item.Click()
                        time.sleep(0.5)
                        print(f"✅ 已通过 Click() 方法选择 '{text}'")
                        return True
                    except Exception as e:
                        print(f"  Click() 方法失败: {e}")
                    
                    # 方法3: 使用 InvokePattern（如果可用）
                    try:
                        invoke_pattern = item.GetInvokePattern()
                        if invoke_pattern:
                            print(f"  使用 InvokePattern 调用 '{text}'...")
                            invoke_pattern.Invoke()
                            time.sleep(0.5)
                            print(f"✅ 已通过 InvokePattern 选择 '{text}'")
                            return True
                    except Exception as e:
                        print(f"  InvokePattern 不可用: {e}")
                    
                    print(f"  ❌ 所有方法都失败，继续查找...")
                    seen_texts.discard(text_norm)
                    continue
            except Exception as e:
                continue
        
        # 打印当前帧的选项（用于调试）
        if current_frame_options:
            print(f"  📋 第 {scrolls_done + 1} 次读取，发现 {new_items_in_frame} 个新项，当前可见选项: {', '.join(current_frame_options[:15])}" + 
                  (f" ... (共{len(current_frame_options)}个)" if len(current_frame_options) > 15 else ""))
        elif scrolls_done > 0:
            print(f"  ⚠️ 第 {scrolls_done + 1} 次读取，未发现新项（可能已到底或滚动未生效）")
        
        # 如果当前帧没有新项且已滚动过，说明可能到底了
        if new_items_in_frame == 0 and scrolls_done > 0:
            print(f"  ℹ️ 已滚动到底，无更多新选项")
            break
        
        # 向下滚动一次（模拟滚轮）- 每滚动一次会出现往下走三个选项
        # 确保鼠标在下拉框区域内
        try:
            if scroll_viewer.Exists(maxSearchSeconds=0.1):
                sv_rect = scroll_viewer.BoundingRectangle
                if sv_rect.width() > 0 and sv_rect.height() > 0:
                    current_x = max(sv_rect.left + 10, min(scroll_x, sv_rect.right - 10))
                    current_y = max(sv_rect.top + 10, min(scroll_y, sv_rect.bottom - 10))
                    user32.SetCursorPos(int(current_x), int(current_y))
                    time.sleep(0.05)
        except:
            pass
        
        user32.mouse_event(0x0800, 0, 0, -120, 0)
        time.sleep(0.6)  # 等待UI更新，确保新选项出现
        scrolls_done += 1
    
    # 打印所有扫描到的选项
    print(f"\n📋 所有扫描到的选项（共 {len(all_scanned_options)} 个）:")
    for i, opt in enumerate(all_scanned_options, 1):
        opt_norm = _normalize_text(opt.split('(')[0])
        match_mark = " ✓" if opt_norm == target_clean else ""
        print(f"  {i}. '{opt}'（标准化: '{opt_norm}'）{match_mark}")
    
    print(f"\n❌ 滚动 {scrolls_done} 次后仍未找到{field_name}: '{target_text}'")
    print(f"   目标文本（原始）: '{target_clean_raw}'")
    print(f"   目标文本（标准化）: '{target_clean}'")
    return False


def select_combobox_by_keyboard(combo, target_text, max_retries=3):
    """通过键盘选择下拉框选项"""
    print(f"⌨️ 选择: '{target_text}'")
    for attempt in range(max_retries):
        try:
            combo.SetFocus()
        except:
            pass
        human_delay(0.6, 1.2)
        
        auto.SendKeys("{F4}")
        human_delay(0.7, 1.3)
        
        clean_text = target_text.split('(')[0].strip()
        pinyin = clean_text[:6]
        auto.SendKeys(pinyin)
        human_delay(0.5, 1.0)
        
        auto.SendKeys("{ENTER}")
        human_delay(0.5, 0.8)
        
        try:
            current_value = combo.GetValuePattern().Value or ""
            # 必须与整项文案一致（标准化后），不能用 in：否则「指甲油」会误判匹配「指甲油底胶及封层胶产品」
            cur_raw = current_value.split("(")[0].strip()
            if _normalize_text(clean_text) == _normalize_text(cur_raw):
                print(f"✅ 选择成功: {current_value}")
                return True
        except:
            pass
    
    print(f"⚠️ 键盘选择失败，尝试滚动查找...")
    return select_dropdown_option_with_scroll(combo, target_text)

