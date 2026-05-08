"""
提取所有类目文本的脚本（新方式）
按照流程图逻辑：先获取所有文本，再用键盘输入方式选择
"""

import time
import json
import ctypes
import uiautomation as auto
from typing import List, Dict, Any
try:
    from pywinauto import Desktop
except:
    pass


def human_delay(min_s=0.5, max_s=1.0):
    """模拟人类操作的随机延迟"""
    import random
    time.sleep(random.uniform(min_s, max_s))


def get_display_text(list_item):
    """提取ListItem的真实显示文本"""
    def dfs(ctrl, depth=0):
        if depth > 4:
            return ""
        try:
            child = ctrl.GetFirstChildControl()
            while child:
                if child.ControlType == auto.ControlType.TextControl:
                    text = child.Name or ""
                    if text and not text.startswith("SHOPEE.") and len(text) < 100:
                        return text.strip()
                result = dfs(child, depth + 1)
                if result:
                    return result
                child = child.GetNextSiblingControl()
        except:
            pass
        return ""
    
    # 优先使用LegacyIAccessible.Name
    try:
        legacy = list_item.LegacyIAccessiblePattern()
        if legacy:
            name = legacy.Name or ""
            if name and not name.startswith("SHOPEE.") and len(name) < 100:
                return name.strip()
    except:
        pass
    
    # 使用ValuePattern
    try:
        value_pattern = list_item.GetValuePattern()
        if value_pattern:
            value = value_pattern.Value or ""
            if value and not value.startswith("SHOPEE.") and len(value) < 100:
                return value.strip()
    except:
        pass
    
    # 递归搜索子控件
    return dfs(list_item)


def get_all_dropdown_options(combo_ctrl) -> List[str]:
    """
    获取下拉框的所有选项文本
    :param combo_ctrl: ComboBox控件
    :return: 选项文本列表
    """
    options = []
    seen_runtime_ids = set()
    
    try:
        # 点击下拉框打开
        combo_ctrl.SetFocus()
        human_delay(0.5, 0.8)
        
        # 打开下拉框
        auto.SendKeys("{F4}")
        human_delay(0.8, 1.2)
        
        # 查找下拉容器
        scroll_viewer = auto.PaneControl(AutomationId="DropDownScrollViewer")
        if not scroll_viewer.Exists(maxSearchSeconds=2):
            # 直接查找ListItem（使用安全遍历方式）
            try:
                child = combo_ctrl.GetFirstChildControl()
                while child:
                    if child.ControlType == auto.ControlType.ListItemControl:
                        try:
                            text = get_display_text(child)
                            if text:
                                options.append(text)
                        except:
                            pass
                    try:
                        child = child.GetNextSiblingControl()
                    except:
                        break
            except:
                pass
            try:
                auto.SendKeys("{Esc}")
            except:
                pass
            return options
        
        # 获取滚动区域
        sv_rect = scroll_viewer.BoundingRectangle
        if sv_rect.width() <= 0 or sv_rect.height() <= 0:
            try:
                auto.SendKeys("{Esc}")
            except:
                pass
            return options
        
        user32 = ctypes.windll.user32
        scroll_x = sv_rect.left + sv_rect.width() // 2
        scroll_y = sv_rect.top + sv_rect.height() // 2
        
        # 将鼠标移动到滚动区域中心
        user32.SetCursorPos(int(scroll_x), int(scroll_y))
        human_delay(0.3, 0.5)
        
        # 滚动到顶部
        for _ in range(5):
            user32.mouse_event(0x0800, 0, 0, 120, 0)  # 向上滚动
            time.sleep(0.2)
        
        human_delay(0.5, 0.8)
        
        # 开始滚动并收集所有选项
        scrolls_done = 0
        max_scrolls = 200  # 最大滚动次数，防止无限循环
        no_new_items_count = 0  # 连续没有新项的计数
        
        while scrolls_done < max_scrolls:
            # 获取当前可见的ListItem（直接遍历，不存储到列表）
            new_items_in_frame = 0
            
            try:
                child = scroll_viewer.GetFirstChildControl()
                while child:
                    if child.ControlType == auto.ControlType.ListItemControl and not child.IsOffscreen:
                        try:
                            runtime_id = None
                            try:
                                runtime_id = tuple(child.GetRuntimeId())
                            except:
                                rect = child.BoundingRectangle
                                if rect.width() <= 0 or rect.height() <= 0:
                                    child = child.GetNextSiblingControl()
                                    continue
                                runtime_id = (rect.left, rect.top, rect.width(), rect.height())
                            
                            if runtime_id in seen_runtime_ids:
                                child = child.GetNextSiblingControl()
                                continue
                            
                            seen_runtime_ids.add(runtime_id)
                            new_items_in_frame += 1
                            
                            text = get_display_text(child).strip()
                            if text and text not in options:
                                options.append(text)
                        except Exception as e:
                            pass
                    
                    try:
                        child = child.GetNextSiblingControl()
                    except:
                        break
            except:
                pass
            
            # 如果没有新项
            if new_items_in_frame == 0:
                no_new_items_count += 1
                # 连续3次没有新项，认为已经到底了
                if no_new_items_count >= 3:
                    break
            else:
                no_new_items_count = 0
            
            # 向下滚动
            user32.mouse_event(0x0800, 0, 0, -120, 0)
            time.sleep(0.3)
            scrolls_done += 1
        
        # 关闭下拉框
        auto.SendKeys("{Esc}")
        human_delay(0.3, 0.5)
        
    except Exception as e:
        try:
            auto.SendKeys("{Esc}")
        except:
            pass
    
    return options


def select_category_by_keyboard(combo, option_text: str) -> bool:
    """
    使用键盘输入方式选择类目选项
    :param combo: ComboBox控件
    :param option_text: 选项文本（完整文本）
    :return: 是否成功选择
    """
    try:
        combo.SetFocus()
        human_delay(0.5, 0.8)
        auto.SendKeys("{F4}")
        human_delay(0.7, 1.0)
        
        # 输入完整文本（不是只输入前几个字符）
        auto.SendKeys(option_text)
        human_delay(0.8, 1.2)
        auto.SendKeys("{Enter}")
        human_delay(1.5, 2.0)  # 等待类目加载
        return True
    except Exception as e:
        return False


def format_option_text(text: str, level: int) -> str:
    """
    格式化选项文本：一级和五级（叶级）不带括号，二至四级带括号
    :param text: 原始文本
    :param level: 层级（1-5）
    :return: 格式化后的文本
    """
    if level == 1 or level == 5:
        # 一级和五级：只取括号前的部分，如果没有括号就取全部
        if '(' in text:
            return text.split('(')[0].strip()
        return text.strip()
    else:
        # 二至四级：保持完整格式（带括号）
        return text.strip()


def process_category_level(window, level: int, parent_key: str, parent_option_text: str, 
                          combo_ids: dict, level_names: list) -> Dict[str, Any]:
    """
    处理某一层级的类目（按照流程图逻辑）
    关键：先获取所有选项，再逐个选择并处理
    注意：不在这里选择父级类目，而是在调用此函数之前已经选择了父级
    :param window: 窗口对象
    :param level: 当前层级（1-5）
    :param parent_key: 父级键名（用于日志）
    :param parent_option_text: 父级选项文本（仅用于日志，不用于选择）
    :param combo_ids: ComboBox ID字典
    :param level_names: 层级名称列表
    :return: 该类目的子级字典
    """
    current_level_name = level_names[level - 1]
    combo_id = combo_ids[current_level_name]
    
    # 步骤1: 查看当前层级控件是否存在
    combo = window.ComboBoxControl(AutomationId=combo_id)
    if not combo.Exists(maxSearchSeconds=2):
        return {}
    
    # 步骤2: 【关键】先获取当前层级所有选项文本（在遍历选择之前）
    options = get_all_dropdown_options(combo)
    
    if not options:
        return {}
    
    # 步骤3: 如果是五级类目（叶级），直接返回所有选项（值为空字符串""）
    if level == 5:
        result = {}
        for option in options:
            formatted_option = format_option_text(option, level)
            result[formatted_option] = ""  # 五级类目的值都是空字符串
        return result
    
    # 步骤4: 按顺序选择当前层级选项，并递归处理下一级
    result = {}
    
    for idx, option in enumerate(options, 1):
        formatted_option = format_option_text(option, level)
        
        # 【关键】先获取选项后再选择：使用键盘输入选择当前选项（输入完整文本）
        select_category_by_keyboard(combo, option)
        
        # 步骤5: 检测是否有下一级类目（控件检测 + Children检测，两者都必须通过）
        next_level = level + 1
        
        # 如果已经是五级类目，肯定没有下一级了
        if level >= 5:
            result[formatted_option] = ""
        else:
            # 重要：先确保所有下拉框都已关闭，等待UI更新
            try:
                auto.SendKeys("{Esc}")  # 确保关闭任何打开的下拉框
            except:
                pass
            human_delay(1.0, 1.5)  # 等待UI更新完成
            
            next_level_name = level_names[next_level - 1]
            next_combo_id = combo_ids[next_level_name]
            
            # 检测是否有下一级类目（按照 ttttt.py 的方式检测控件）
            next_combo = None
            has_list_items = False
            
            try:
                # 按照 ttttt.py 的方式获取控件
                desktop = Desktop(backend="uia")
                pywinauto_window = desktop.window(title="类目修改")
                if pywinauto_window.exists():
                    combo_wrapper = pywinauto_window.child_window(auto_id=next_combo_id, control_type="ComboBox").wrapper_object()
                    # 保留 uiautomation 对象用于后续操作
                    next_combo = window.ComboBoxControl(AutomationId=next_combo_id)
                    
                    # 【关键修复】在检测Children之前，先刷新下一级控件的状态
                    # 问题：当处理完三级类目回到二级类目时，如果此时的二级类目没有三级类目，
                    # 则三级类目的Children的ListItemControl并不会重置为空，而是保留上次展开时的ListItemControl
                    # 解决：先尝试展开并关闭下拉框，以刷新控件状态
                    try:
                        rect = next_combo.BoundingRectangle
                        if rect.width() > 0 and rect.height() > 0:
                            auto.Click(rect.left + rect.width() // 2, rect.top + rect.height() // 2)
                            human_delay(0.3, 0.5)
                            auto.SendKeys("{F4}")
                            human_delay(0.5, 0.8)
                            auto.SendKeys("{Esc}")
                            human_delay(0.5, 0.8)
                    except:
                        pass
                    
                    # 检测控件的Children中是否有ListItemControl
                    try:
                        child = next_combo.GetFirstChildControl()
                        while child:
                            try:
                                if child.ControlType == auto.ControlType.ListItemControl:
                                    has_list_items = True
                                    break
                            except:
                                pass
                            try:
                                child = child.GetNextSiblingControl()
                            except:
                                break
                    except:
                        pass
            except:
                pass
            
            has_next_level = has_list_items
            
            if not has_next_level:
                result[formatted_option] = ""  # 没有下一级，值为空字符串
            else:
                child_result = process_category_level(
                    window, next_level, formatted_option, option, combo_ids, level_names
                )
                if child_result:
                    result[formatted_option] = child_result
                else:
                    result[formatted_option] = ""  # 没有子级，值为空字符串
    
    return result


def main():
    """主函数"""
    print("=" * 60)
    print("🚀 开始提取所有类目文本（新方式）")
    print("=" * 60)
    
    # 查找窗口
    window = auto.WindowControl(Name="类目修改", searchDepth=1)
    if not window.Exists(maxSearchSeconds=5):
        print("❌ 未找到'类目修改'窗口，请先打开窗口")
        return
    
    window.SetActive()
    human_delay(1.0, 1.5)
    print("✅ 已找到并激活'类目修改'窗口")
    
    # 输出文件设置为"文本.json"
    output_file = "文本.json"
    existing_categories = {}
    skip_categories = set()
    
    # 步骤2: 读取现有的"文本.json"文件，没有则创建
    try:
        with open(output_file, "r", encoding="utf-8") as f:
            existing_categories = json.load(f)
            skip_categories = set(existing_categories.keys())
            print(f"\n📂 已读取现有文件，将跳过已存在的类目: {skip_categories}")
    except:
        print(f"\n📂 未找到现有文件，将创建新文件")
        # 创建空文件
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
    
    combo_ids = {
        "一级": "OneTypeSkinComboBox",
        "二级": "TwoTypeSkinComboBox",
        "三级": "ThreeSkinComboBox",
        "四级": "FourSkinComboBox",
        "五级": "FiveSkinComboBox",
    }
    level_names = ["一级", "二级", "三级", "四级", "五级"]
    
    # 步骤3: 获取一级类目的所有选项，输出到文本.json文件中，按照格式
    print(f"\n{'=' * 60}")
    print(f"📋 步骤3: 获取一级类目所有文本")
    print(f"{'=' * 60}")
    
    combo = window.ComboBoxControl(AutomationId=combo_ids["一级"])
    if not combo.Exists(maxSearchSeconds=3):
        print("❌ 未找到一级类目控件")
        return
    
    all_level1_options = get_all_dropdown_options(combo)
    print(f"\n✅ 共找到 {len(all_level1_options)} 个一级类目")
    
    # 创建初始JSON结构，每个一级类目作为键，值为空字符串（如果不存在）
    categories = existing_categories.copy()
    for level1_option in all_level1_options:
        formatted_level1 = format_option_text(level1_option, 1)
        if formatted_level1 not in categories:
            categories[formatted_level1] = ""
    
    # 保存初始结构到JSON文件
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(categories, f, ensure_ascii=False, indent=4)
    print(f"✅ 已创建初始JSON结构并保存到 {output_file}")
    
    # 步骤4: 遍历每个一级类目
    for idx, level1_option in enumerate(all_level1_options, 1):
        formatted_level1 = format_option_text(level1_option, 1)
        
        if formatted_level1 in skip_categories:
            print(f"\n⏭️  [{idx}/{len(all_level1_options)}] 跳过已存在的类目: {formatted_level1}")
            continue
        
        print(f"\n{'=' * 60}")
        print(f"📂 [{idx}/{len(all_level1_options)}] 开始处理一级类目: {formatted_level1}")
        print(f"{'=' * 60}")
        
        # 选择该一级类目
        select_category_by_keyboard(combo, level1_option)
        
        # 检测是否有二级类目，如果有则递归处理
        child_result = process_category_level(
            window, level=2, parent_key=formatted_level1, 
            parent_option_text=level1_option, combo_ids=combo_ids, level_names=level_names
        )
        
        # 更新结果（如果没有子级，保持为空字符串）
        if child_result:
            categories[formatted_level1] = child_result
        else:
            categories[formatted_level1] = ""
        
        # 步骤5: 每处理完一个一级类目就保存到"文本.json"
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(categories, f, ensure_ascii=False, indent=4)
        print(f"✅ 已保存一级类目 '{formatted_level1}' 的处理结果")
    
    print("\n" + "=" * 60)
    print(f"✅ 类目提取完成！已保存到: {output_file}")
    print("=" * 60)
    
    # 步骤6: 输出统计信息
    def count_categories(cat_dict, level=1):
        count = 0
        for key, value in cat_dict.items():
            count += 1
            if isinstance(value, dict) and value:
                count += count_categories(value, level + 1)
        return count
    
    total_count = count_categories(categories)
    print(f"\n📊 统计信息:")
    print(f"   总类目数: {total_count}")
    print(f"   一级类目数: {len(categories)}")
    print(f"   跳过类目数: {len(skip_categories)}")


if __name__ == "__main__":
    main()
