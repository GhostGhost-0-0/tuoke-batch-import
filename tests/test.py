import uiautomation as auto
import time
import json
import ctypes
import random
import os

# === 配置文件名（从文件读取所有配置）===
CONFIG_FILE = "属性配置.json"  # 配置文件路径

def load_config(config_file=CONFIG_FILE):
    """
    从JSON文件加载配置（包括类目和属性）
    配置文件格式示例：
    {
        "类目": {
            "一级": "美妆保健(美妆保健)",
            "二级": "手足保养与美甲(手足保养与美甲)",
            "三级": "手部保养(手部保养)",
            "四级": "手膜"  // 留空则跳过
        },
        "属性": {
            "品牌": "chloe",
            "保质期": "24个月",
            "颜色": "红色",
            ...
        }
    }
    """
    if not os.path.exists(config_file):
        print(f"⚠️ 配置文件 '{config_file}' 不存在，使用默认配置")
        return {
            "类目": {
                "一级": "美妆保健(美妆保健)",
                "二级": "手足保养与美甲(手足保养与美甲)",
                "三级": "手部保养(手部保养)",
                "四级": "手膜"
            },
            "属性": {
                "品牌": "chloe",
                "保质期": "24个月"
            }
        }
    
    try:
        with open(config_file, "r", encoding="utf-8") as f:
            config = json.load(f)
        
        # 兼容旧格式（如果直接是属性配置）
        if "类目" not in config and "属性" not in config:
            print(f"ℹ️ 检测到旧格式配置文件，转换为新格式")
            config = {
                "类目": {
                    "一级": "美妆保健(美妆保健)",
                    "二级": "手足保养与美甲(手足保养与美甲)",
                    "三级": "手部保养(手部保养)",
                    "四级": "手膜"
                },
                "属性": config
            }
        
        # 确保类目和属性都存在
        if "类目" not in config:
            config["类目"] = {
                "一级": "美妆保健(美妆保健)",
                "二级": "手足保养与美甲(手足保养与美甲)",
                "三级": "手部保养(手部保养)",
                "四级": "手膜"
            }
        if "属性" not in config:
            config["属性"] = {}
        
        print(f"✅ 已从 '{config_file}' 加载配置")
        print(f"   类目: {config.get('类目', {})}")
        print(f"   属性: {config.get('属性', {})}")
        return config
    except Exception as e:
        print(f"❌ 读取配置文件失败: {e}")
        return {
            "类目": {
                "一级": "美妆保健(美妆保健)",
                "二级": "手足保养与美甲(手足保养与美甲)",
                "三级": "手部保养(手部保养)",
                "四级": "手膜"
            },
            "属性": {}
        }

def human_delay(min_s=0.8, max_s=2.5):
    time.sleep(random.uniform(min_s, max_s))

def mouse_click(x, y):
    if x < 50 or y < 50 or x > 3000 or y > 2000:
        return False
    user32 = ctypes.windll.user32
    user32.SetCursorPos(int(x), int(y))
    time.sleep(0.05)
    user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFT DOWN
    time.sleep(0.03)
    user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFT UP
    return True

def select_combobox_by_keyboard(combo, target_text, max_retries=3):
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
        print(f"  → 输入拼音: '{pinyin}' (第 {attempt+1} 次)")
        auto.SendKeys(pinyin)
        human_delay(1.0, 1.8)

        auto.SendKeys("{Enter}")
        human_delay(0.9, 1.5)

        try:
            current = combo.GetValuePattern().Value or ""
            if target_text in current:
                print(f"✅ 成功: {current}")
                return True
        except Exception as e:
            print(f"  ⚠️ 读取值失败: {e}")

        if attempt < max_retries - 1:
            human_delay(1.2, 2.0)
    print(f"❌ 失败: '{target_text}'")
    return False

# --- 提取 ListItem 的真实显示文本 ---
def get_display_text(list_item):
    def dfs(ctrl, depth=0):
        if depth > 4:
            return ""
        try:
            child = ctrl.GetFirstChildControl()
            while child:
                if child.ControlType == auto.ControlType.TextControl:
                    name = (child.Name or "").strip()
                    if name and len(name) <= 100 and not name.startswith("SHOPEE."):
                        return name
                res = dfs(child, depth + 1)
                if res:
                    return res
                child = child.GetNextSiblingControl()
        except:
            pass
        return ""
    return dfs(list_item).strip()

# --- 【核心】下拉框：滚动条滚动 + 匹配 + 点击（通用函数）---
def select_dropdown_option_with_scroll(combo_ctrl, target_text, field_name="选项"):
    """
    在下拉框中滚动查找并点击目标选项
    :param combo_ctrl: 下拉框控件
    :param target_text: 目标文本
    :param field_name: 字段名称（用于日志显示）
    """
    print(f"\n🔍 在{field_name}下拉框中查找并点击: '{target_text}'（使用滚动条）")

    # === 1. 展开下拉框 ===
    dropdown_button = None
    child = combo_ctrl.GetFirstChildControl()
    while child:
        if child.ControlType == auto.ControlType.ButtonControl and not child.IsOffscreen:
            dropdown_button = child
            break
        child = child.GetNextSiblingControl()

    if dropdown_button:
        try:
            dropdown_button.Click()
        except:
            r = dropdown_button.BoundingRectangle
            if r.width() > 0 and r.height() > 0:
                mouse_click(r.left + r.width() // 2, r.top + r.height() // 2)
    else:
        rect = combo_ctrl.BoundingRectangle
        if rect.width() > 30:
            cx = rect.right - 15
            cy = rect.top + rect.height() // 2
            if cx > 0 and cy > 0:
                mouse_click(cx, cy)

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

    user32 = ctypes.windll.user32
    scroll_x = sv_rect.left + sv_rect.width() // 2
    scroll_y = sv_rect.top + min(60, sv_rect.height() // 2)
    scroll_x = max(50, min(scroll_x, 2500))
    scroll_y = max(50, min(scroll_y, 1800))
    user32.SetCursorPos(int(scroll_x), int(scroll_y))
    time.sleep(0.1)

    seen_runtime_ids = set()  # 使用 RuntimeId 去重（最可靠）
    target_clean = target_text.strip().lower()
    max_scrolls = 30
    scrolls_done = 0
    all_scanned_options = []  # 记录所有扫描到的选项

    while scrolls_done <= max_scrolls:
        # 获取当前所有可见 ListItem
        items = []
        try:
            child = scroll_viewer.GetFirstChildControl()
            while child:
                if child.ControlType == auto.ControlType.ListItemControl and not child.IsOffscreen:
                    items.append(child)
                child = child.GetNextSiblingControl()
        except:
            items = []

        if not items and scrolls_done == 0:
            print("⚠️ 初始无任何选项")
            break

        # 检查当前帧是否包含目标
        current_frame_options = []  # 当前帧的选项
        new_items_in_frame = 0  # 当前帧中的新项数量
        
        for item in items:
            try:
                # 使用 RuntimeId 去重（最可靠，不会因为坐标变化而误判）
                runtime_id = None
                try:
                    runtime_id = tuple(item.GetRuntimeId())
                except:
                    # 如果获取 RuntimeId 失败，使用坐标作为备用
                    rect = item.BoundingRectangle
                    if rect.width() <= 0 or rect.height() <= 0:
                        continue
                    runtime_id = (rect.left, rect.top, rect.width(), rect.height())
                
                if runtime_id in seen_runtime_ids:
                    continue  # 已处理过的项，跳过
                
                seen_runtime_ids.add(runtime_id)
                new_items_in_frame += 1

                text = get_display_text(item).strip()
                if not text:
                    text = (item.Name or "").strip()
                    if text.startswith("SHOPEE.") or len(text) > 100:
                        text = ""

                # 记录选项文本
                if text:
                    current_frame_options.append(text)
                    if text not in all_scanned_options:
                        all_scanned_options.append(text)

                if text.strip().lower() == target_clean:
                    rect = item.BoundingRectangle
                    cx, cy = rect.left + rect.width() // 2, rect.top + rect.height() // 2
                    if 50 < cx < 2500 and 50 < cy < 1800:
                        print(f"✅ 找到{field_name} '{text}'，正在点击...")
                        # 先尝试直接点击控件
                        try:
                            item.Click()
                            print(f"✅ 已通过 Click() 方法点击{field_name} '{text}'")
                            time.sleep(0.5)
                            return True
                        except:
                            # 如果直接点击失败，使用坐标点击
                            if mouse_click(cx, cy):
                                print(f"✅ 已通过坐标点击{field_name} '{text}'")
                                time.sleep(0.5)
                                return True
                            else:
                                print(f"⚠️ 点击{field_name} '{text}' 失败")
                                return False
            except Exception as e:
                continue

        # 打印当前帧的选项（用于调试）
        if current_frame_options:
            print(f"  📋 第 {scrolls_done + 1} 次滚动，发现 {new_items_in_frame} 个新项，当前可见选项: {', '.join(current_frame_options[:10])}" + 
                  (f" ... (共{len(current_frame_options)}个)" if len(current_frame_options) > 10 else ""))
        elif scrolls_done > 0:
            print(f"  ⚠️ 第 {scrolls_done + 1} 次滚动，未发现新项（可能已到底或滚动未生效）")

        # 如果当前帧没有新项且已滚动过，说明可能到底了
        if new_items_in_frame == 0 and scrolls_done > 0:
            print(f"  ℹ️ 已滚动到底，无更多新选项")
            break

        # 向下滚动一次（模拟滚轮）
        user32.mouse_event(0x0800, 0, 0, -120, 0)
        time.sleep(0.5)  # 增加等待时间，确保UI更新
        scrolls_done += 1

    # 打印所有扫描到的选项
    print(f"\n📋 所有扫描到的选项（共 {len(all_scanned_options)} 个）:")
    for i, opt in enumerate(all_scanned_options, 1):
        match_mark = " ✓" if opt.strip().lower() == target_clean else ""
        print(f"  {i}. '{opt}'{match_mark}")
    
    print(f"\n❌ 滚动 {scrolls_done} 次后仍未找到{field_name}: '{target_text}'")
    print(f"   目标文本（小写）: '{target_clean}'")
    return False

# --- 辅助函数 ---
def get_all_descendants(ctrl, depth=0, max_depth=8):
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

def find_new_controls(window, old_snapshot):
    new_snapshot = take_snapshot(window)
    diff = new_snapshot - old_snapshot
    candidates = []
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
            if fp in diff:
                candidates.append(ctrl)
        except:
            continue
    return candidates

def pair_label_with_input(window):
    labels = []
    inputs = []
    all_controls = get_all_descendants(window)
    keywords = ["品牌", "颜色", "材质", "尺寸", "风格", "系列", "适用", "图案", "闭合", "肩带", "类别", "型号", "季节", "袖长", "保质期"]
    
    for ctrl in all_controls:
        name = (ctrl.Name or "").strip()
        if (ctrl.ControlType == auto.ControlType.TextControl 
            and any(k in name for k in keywords) 
            and 1 <= len(name) <= 8):
            labels.append(ctrl)
        elif (ctrl.ControlType in [auto.ControlType.EditControl, auto.ControlType.ComboBoxControl]
              and not ctrl.IsOffscreen
              and ctrl.IsEnabled):
            inputs.append(ctrl)

    pairs = []
    for label in labels:
        label_rect = label.BoundingRectangle
        best_input = None
        min_dx = float('inf')
        for inp in inputs:
            inp_rect = inp.BoundingRectangle
            if inp_rect.width() <= 0 or inp_rect.height() <= 0:
                continue
            if (inp_rect.top - 10 < label_rect.bottom and 
                inp_rect.bottom + 10 > label_rect.top):
                dx = inp_rect.left - label_rect.right
                if 0 < dx < 350 and dx < min_dx:
                    min_dx = dx
                    best_input = inp
        if best_input:
            pairs.append((label, best_input))
    return pairs

def handle_dialog_popup(parent_window, timeout=5):
    """
    检测并处理弹窗对话框，自动点击"确定"按钮
    :param parent_window: 父窗口对象
    :param timeout: 等待弹窗出现的超时时间（秒）
    :return: 是否成功处理弹窗
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

# --- 主流程 ---
def main():
    print("🎯 开始流程：类目 → 加载属性 → 品牌滚动匹配点击 + 保质期选择\n")
    total_start = time.perf_counter()

    window = auto.WindowControl(Name="类目修改", searchDepth=1)
    if not window.Exists(maxSearchSeconds=5):
        print("❌ 未找到 '类目修改' 窗口")
        return
    window.SetActive()
    human_delay(1.0, 1.8)

    # === 从配置文件加载类目配置 ===
    config = load_config()
    category_config = config.get("类目", {})
    
    # 获取类目目标值（从配置文件读取，如果没有则使用默认值）
    PRIMARY_TARGET = category_config.get("一级", "美妆保健(美妆保健)")
    SECONDARY_TARGET = category_config.get("二级", "手足保养与美甲(手足保养与美甲)")
    TERTIARY_TARGET = category_config.get("三级", "手部保养(手部保养)")
    FOURTH_TARGET = category_config.get("四级", "手膜")
    
    print(f"\n📋 类目配置:")
    print(f"   一级: {PRIMARY_TARGET}")
    print(f"   二级: {SECONDARY_TARGET}")
    print(f"   三级: {TERTIARY_TARGET}")
    print(f"   四级: {FOURTH_TARGET if FOURTH_TARGET.strip() else '(跳过)'}")

    primary_combo   = window.ComboBoxControl(AutomationId="OneTypeSkinComboBox")
    secondary_combo = window.ComboBoxControl(AutomationId="TwoTypeSkinComboBox")
    tertiary_combo  = window.ComboBoxControl(AutomationId="ThreeSkinComboBox")
    fourth_combo    = window.ComboBoxControl(AutomationId="FourSkinComboBox")

    for name, ctrl in [("一级", primary_combo), ("二级", secondary_combo), ("三级", tertiary_combo)]:
        if not ctrl.Exists(maxSearchSeconds=3):
            print(f"❌ {name}类目控件不存在")
            return

    steps = [
        ("一级", primary_combo, PRIMARY_TARGET),
        ("二级", secondary_combo, SECONDARY_TARGET),
        ("三级", tertiary_combo, TERTIARY_TARGET),
    ]
    if fourth_combo.Exists(maxSearchSeconds=1) and FOURTH_TARGET.strip():
        steps.append(("四级", fourth_combo, FOURTH_TARGET))

    for level_name, combo, target in steps:
        print(f"\n--- 选择{level_name}类目: {target} ---")
        if not select_combobox_by_keyboard(combo, target):
            return
        human_delay(2.0, 3.0)

    print("\n📸 拍摄加载前 UI 快照...")
    before_snapshot = take_snapshot(window)

    load_button = window.ButtonControl(AutomationId="UpdateSkinButton") or window.ButtonControl(Name="加载类目属性")
    if load_button.Exists():
        try:
            load_button.Click()
        except:
            r = load_button.BoundingRectangle
            mouse_click(r.left + r.width() // 2, r.top + r.height() // 2)
        print("✅ 已点击“加载类目属性”")
        human_delay(1.0, 1.5)  # 先等待一下，让弹窗有时间出现
        
        # === 检测并处理弹窗 ===
        handle_dialog_popup(window, timeout=3)
        
        human_delay(2.0, 3.0)  # 继续等待属性加载完成
    else:
        print("❌ 未找到“加载类目属性”按钮")
        return

    print("🔍 对比快照，寻找新增控件...")
    new_controls = find_new_controls(window, before_snapshot)
    print(f"   新增控件数量: {len(new_controls)}")

    pairs = pair_label_with_input(window)
    if not pairs:
        print("\n⚠️ 未识别到任何属性字段。")
        return

    print(f"\n✅ 成功识别 {len(pairs)} 个属性字段:")
    results = []
    for label, inp in pairs:
        name = label.Name.strip()
        try:
            val = inp.GetValuePattern().Value or ""
        except:
            val = inp.Name or "N/A"
        print(f"  • {name}: {val}")
        results.append({"name": name, "value": val})

    # === 从配置文件加载属性配置 ===
    attribute_config = config.get("属性", {})
    
    if not attribute_config:
        print("\n⚠️ 未加载到属性配置，跳过属性设置")
    else:
        # === 根据配置文件动态处理属性字段 ===
        print(f"\n📋 开始处理 {len(attribute_config)} 个属性配置...")
        
        for attr_name, attr_value in attribute_config.items():
            print(f"\n🔄 处理属性: '{attr_name}' = '{attr_value}'")
            handled = False
            
            # 在识别的字段中查找匹配的属性
            for label, inp in pairs:
                # 模糊匹配属性名（支持部分匹配）
                if attr_name in label.Name or label.Name in attr_name:
                    print(f"  ✓ 找到匹配的属性字段: '{label.Name}'")
                    
                    if inp.ControlType == auto.ControlType.ComboBoxControl:
                        # 下拉框：滚动查找并点击
                        if select_dropdown_option_with_scroll(inp, attr_value, field_name=attr_name):
                            handled = True
                            print(f"  ✅ 已设置 {attr_name} 为: {attr_value}")
                        else:
                            print(f"  ❌ 设置 {attr_name} 失败，未找到选项: {attr_value}")
                    elif inp.ControlType == auto.ControlType.EditControl:
                        # 输入框：直接设置值
                        try:
                            value_pattern = inp.GetValuePattern()
                            if value_pattern:
                                value_pattern.SetValue(attr_value)
                                handled = True
                                print(f"  ✅ 已设置 {attr_name} 为: {attr_value}")
                            else:
                                # 如果 ValuePattern 不可用，尝试键盘输入
                                inp.SetFocus()
                                time.sleep(0.2)
                                auto.SendKeys("^a")  # Ctrl+A 全选
                                time.sleep(0.1)
                                auto.SendKeys(attr_value)
                                handled = True
                                print(f"  ✅ 已通过键盘输入设置 {attr_name} 为: {attr_value}")
                        except Exception as e:
                            print(f"  ❌ 设置 {attr_name} 失败: {e}")
                    else:
                        print(f"  ⚠️ {attr_name} 字段类型不支持: {inp.ControlType}")
                    
                    break  # 找到匹配的字段后退出循环
            
            if not handled:
                print(f"  ⚠️ 未找到匹配的属性字段: '{attr_name}'")
            
            # 每个属性处理完后稍作延迟
            if handled:
                human_delay(0.5, 1.0)
        
        # === 所有属性处理完成后，点击确认按钮 ===
        print(f"\n✅ 所有属性处理完成，准备点击'确认'按钮...")
        confirm_button = window.ButtonControl(AutomationId="skinButton1") or window.ButtonControl(Name="确认")
        
        if confirm_button.Exists(maxSearchSeconds=3):
            try:
                # 先尝试直接点击
                confirm_button.Click()
                print("✅ 已点击'确认'按钮")
                human_delay(1.0, 1.5)
            except:
                # 如果Click失败，使用坐标点击
                try:
                    r = confirm_button.BoundingRectangle
                    if r.width() > 0 and r.height() > 0:
                        cx = r.left + r.width() // 2
                        cy = r.top + r.height() // 2
                        if mouse_click(cx, cy, window=window):
                            print("✅ 已通过坐标点击'确认'按钮")
                            human_delay(1.0, 1.5)
                        else:
                            print("❌ 点击'确认'按钮失败")
                except Exception as e:
                    print(f"❌ 点击'确认'按钮失败: {e}")
        else:
            print("⚠️ 未找到'确认'按钮（AutomationId='skinButton1' 或 Name='确认'）")

    with open("属性识别结果.json", "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print("\n💾 属性识别结果已保存到 '属性识别结果.json'")

    total_end = time.perf_counter()
    print(f"\n🏁 全程耗时: {total_end - total_start:.2f} 秒")
    print("✅ 四级类目 + 品牌滚动匹配点击 + 保质期选择 + 确认 完成！")

if __name__ == "__main__":
    main()