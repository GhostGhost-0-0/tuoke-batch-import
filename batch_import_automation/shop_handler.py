"""
店铺选择处理模块
"""

import re
import time
from pywinauto import Application, mouse, Desktop


class ShopHandler:
    """店铺选择处理类"""

    def __init__(self, window=None):
        """
        初始化店铺处理器
        :param window: 窗口控件对象（保留以兼容现有代码，但实际不使用）
        """
        self.window = window
        self._main_app = None
        self._main_win = None

    def _get_main_window(self, force_reconnect: bool = False):
        """
        获取并聚焦「批量导入」窗口。
        优先复用已有连接，失效时自动重连。
        :param force_reconnect: 是否强制重连
        :return: pywinauto window 对象或 None
        """
        if force_reconnect:
            self._main_app = None
            self._main_win = None

        # 先尝试复用已有窗口连接，避免每次 connect 带来的额外耗时
        if self._main_win is not None:
            try:
                self._main_win.wait("visible", timeout=1)
                self._main_win.set_focus()
                time.sleep(0.2)
                print("♻️ 复用已连接的 '批量导入' 窗口")
                return self._main_win
            except Exception:
                print("⚠️ 已缓存窗口失效，准备重新连接")
                self._main_app = None
                self._main_win = None

        # 兜底：重新连接
        try:
            print("🔌 重新连接 '批量导入' 窗口...")
            self._main_app = Application(backend="uia").connect(title="批量导入", timeout=10)
            self._main_win = self._main_app.window(title="批量导入")
            self._main_win.wait("visible", timeout=10)
            self._main_win.set_focus()
            time.sleep(0.3)
            print("✅ 已重新连接并激活 '批量导入' 窗口")
            return self._main_win
        except Exception as e:
            print(f"❌ 无法连接主窗口: {e}")
            self._main_app = None
            self._main_win = None
            return None

    @staticmethod
    def rects_overlap_vertically(r1, r2, threshold=15):
        """判断两个矩形是否垂直重叠（参考 uggiuyguyi.py）"""
        return not (r1.bottom < r2.top - threshold or r1.top > r2.bottom + threshold)
    
    def toggle_shop_by_name(self, popup, name):
        """切换指定店铺（不判断状态，直接操作，参考 uggiuyguyi.py）"""
        list_items = popup.descendants(control_type="ListItem")
        for item in list_items:
            item_name = getattr(item.element_info, 'name', '').strip()
            if item_name == name:
                print(f"🔄 切换店铺: {name}")
                try:
                    item.select()
                    time.sleep(0.2)
                    return True
                except Exception as e1:
                    cbs = item.descendants(control_type="CheckBox")
                    if cbs:
                        try:
                            cbs[0].toggle()
                            time.sleep(0.2)
                            return True
                        except Exception as e2:
                            print(f"⚠️ CheckBox.toggle() 失败: {e2}")
                    else:
                        print(f"❌ 未找到 CheckBox 子控件: {name}")
                break
        else:
            print(f"❌ 未找到店铺: {name}")
            return False
        return False

    def _get_popup_shop_items(self, popup):
        """获取弹窗中的店铺项映射：标准化名称 -> 原始名称"""
        items = popup.descendants(control_type="ListItem")
        name_map = {}
        for item in items:
            raw_name = getattr(item.element_info, "name", "").strip()
            if raw_name and "(已封)" not in raw_name:
                name_map[raw_name] = raw_name
        return name_map

    def _toggle_shop_by_keyword(self, popup, keyword):
        """
        按关键字模糊匹配并切换店铺。
        返回本次成功切换数量。
        """
        kw = (keyword or "").strip()
        if not kw:
            return 0

        matched_names = []
        seen = set()
        for item in popup.descendants(control_type="ListItem"):
            item_name = getattr(item.element_info, "name", "").strip()
            if (
                item_name
                and "(已封)" not in item_name
                and kw in item_name
                and item_name not in seen
            ):
                matched_names.append(item_name)
                seen.add(item_name)

        if not matched_names:
            print(f"❌ 未找到包含关键字的店铺: {kw}")
            return 0

        print(f"🔎 关键字 '{kw}' 模糊匹配到 {len(matched_names)} 个店铺")
        done = 0
        for item_name in matched_names:
            if self.toggle_shop_by_name(popup, item_name):
                done += 1
        return done

    def _toggle_shop_by_rule(self, popup, token, popup_shop_name_map):
        """
        根据输入规则切换店铺：
        - 若 token 可精确命中弹窗店铺名：精准匹配
        - 否则按 token 进行模糊匹配（用于只填类目）
        """
        key = (token or "").strip()
        if not key:
            return 0

        if "(已封)" in key:
            print(f"⏭️ 跳过已封店铺: {key}")
            return 0

        if key in popup_shop_name_map:
            print(f"🎯 精准匹配店铺: {key}")
            return 1 if self.toggle_shop_by_name(popup, key) else 0

        print(f"🧩 未命中完整店铺名，按类目关键字模糊匹配: {key}")
        return self._toggle_shop_by_keyword(popup, key)
    
    def select_shops(self, shop_names: str, old_shop: str = "", is_site_changed: bool = False, select_all_only: bool = False) -> bool:
        """
        选择店铺（使用 pywinauto 方式实现，参考 uggiuyguyi.py）
        :param shop_names: 店铺名称字符串，可能包含多个店铺（用逗号、分号等分隔）
        :param old_shop: 旧店铺名称（用于取消勾选）
        :param is_site_changed: 是否是站点切换（只在站点切换时取消全部勾选）
        :param select_all_only: 是否只执行"全选"操作
        :return: 是否成功选择所有店铺
        """
        print("\n🏪 开始选择店铺...")
        
        # 解析店铺名称（支持多种分隔符：逗号、分号、空格等）
        shop_list_names = re.split(r'[,，;；\n\r]+', str(shop_names))
        shop_list_names = [name.strip() for name in shop_list_names if name.strip()]
        
        if not shop_list_names:
            print("  ⚠️ 未提供店铺名称")
            return False
        
        print(f"  📋 需要选择的店铺数量: {len(shop_list_names)}")
        for i, shop_name in enumerate(shop_list_names, 1):
            print(f"    {i}. {shop_name}")
        
        # 解析旧店铺名称
        old_shop_list = []
        if old_shop:
            old_shop_list = re.split(r'[,，;；\n\r]+', str(old_shop))
            old_shop_list = [name.strip() for name in old_shop_list if name.strip()]
            if old_shop_list:
                print(f"  📋 需要取消勾选的旧店铺数量: {len(old_shop_list)}")
                for i, shop_name in enumerate(old_shop_list, 1):
                    print(f"    {i}. {shop_name}")
        
        # === 步骤1: 获取主窗口「批量导入」===
        main_win = self._get_main_window()
        if not main_win:
            return False

        # === 步骤2: 定位"店铺："标签 ===
        shop_label = None
        for elem in main_win.descendants(control_type="Text"):
            try:
                text = elem.window_text().strip()
                if "店铺" in text:
                    shop_label = elem
                    print(f"✅ 找到店铺标签: '{text}'")
                    break
            except:
                continue

        if not shop_label:
            print("❌ 未找到 '店铺：' 标签，无法展开下拉框")
            return False

        shop_rect = shop_label.rectangle()

        # === 步骤3: 查找与"店铺："邻近的 List 控件 ===
        all_lists = main_win.descendants(control_type="List")
        target_list = None

        if all_lists:
            for lst in all_lists:
                try:
                    lst_rect = lst.rectangle()
                    # 条件：在右侧、水平距离合理（<300px）、垂直对齐
                    if (lst_rect.left > shop_rect.right and
                        lst_rect.left - shop_rect.right < 300 and
                        self.rects_overlap_vertically(shop_rect, lst_rect)):
                        target_list = lst
                        break
                except:
                    continue

        # === 步骤4: 计算点击位置 ===
        if target_list:
            lst_rect = target_list.rectangle()
            click_x = lst_rect.right - 10
            click_y = lst_rect.top + lst_rect.height() // 2
            print(f"🖱️ 智能定位点击: ({click_x}, {click_y}) — 基于 List 控件")
        else:
            # 回退方案：基于"店铺："位置 + 经验偏移
            click_x = shop_rect.right + 190
            click_y = shop_rect.top + shop_rect.height() // 2
            print(f"⚠️ 未找到 List，回退点击: ({click_x}, {click_y})")

        mouse.click(coords=(click_x, click_y))
        time.sleep(0.6)

        # === 步骤5: 查找弹出窗口 ===
        desktop = Desktop(backend="uia")
        popup = None
        for attempt in range(12):
            try:
                for win in desktop.windows():
                    try:
                        if win.class_name() == "Popup":
                            popup = win
                            break
                        checkboxes = win.descendants(control_type="CheckBox", depth=5)
                        if checkboxes:
                            for cb in checkboxes:
                                for child in cb.children():
                                    name = getattr(child.element_info, 'name', '')
                                    if name and "全部" in name:
                                        popup = win
                                        break
                                if popup:
                                    break
                        if popup:
                            break
                    except:
                        continue
                if popup:
                    print("✅ 找到弹出窗口")
                    break
            except Exception as e:
                print(f"⚠️ 扫描异常: {e}")
            time.sleep(0.5)

        if not popup:
            print("❌ 未找到弹出窗口")
            print("所有顶层窗口:")
            for w in desktop.windows():
                try:
                    title = w.window_text().strip()
                    cls = w.class_name()
                    print(f" - Title: {repr(title)}, Class: {cls}")
                except:
                    pass
            return False

        win_rect = main_win.rectangle()

        # === 步骤6: 取消"全部"勾选 ===
        checkboxes = popup.descendants(control_type="CheckBox")
        all_checkbox = None
        for cb in checkboxes:
            for child in cb.children():
                if getattr(child.element_info, 'control_type', '') == 'Text':
                    text_content = getattr(child.element_info, 'name', '').strip()
                    if text_content == "全部":
                        all_checkbox = cb
                        break
            if all_checkbox:
                break

        # 只在站点切换时取消全部勾选
        if is_site_changed and all_checkbox:
            try:
                if hasattr(all_checkbox, 'iface_toggle') and all_checkbox.iface_toggle:
                    state = all_checkbox.iface_toggle.CurrentToggleState
                    if state == 1:
                        print("\n🔄 站点切换，正在取消全部勾选...")
                        all_checkbox.click_input()
                        time.sleep(0.6)
                    else:
                        print("ℹ️ '全部' 未勾选，跳过取消操作")
                else:
                    all_checkbox.toggle()
                    time.sleep(0.6)
            except Exception as e:
                print(f"⚠️ 取消 '全部' 失败: {e}")
        elif is_site_changed and not all_checkbox:
            print("❌ 未找到 '全部' CheckBox")

        # 预取一次弹窗店铺名，供“精确/模糊”规则判断
        popup_shop_name_map = self._get_popup_shop_items(popup)

        # === 步骤7: 取消旧目标店铺 ===
        if old_shop_list and not is_site_changed:
            print("\n🔄 正在取消旧目标店铺...")
            old_done = 0
            for name in old_shop_list:
                old_done += self._toggle_shop_by_rule(popup, name, popup_shop_name_map)

            print(f"🎉 旧目标取消: {old_done}/{len(old_shop_list)}")

        # === 步骤8: 勾选新目标店铺 ===
        new_done = 0
        
        if select_all_only:
            # 只执行"全选"操作
            print("\n🆕 正在执行全选操作...")
            if all_checkbox:
                try:
                    if hasattr(all_checkbox, 'iface_toggle') and all_checkbox.iface_toggle:
                        state = all_checkbox.iface_toggle.CurrentToggleState
                        if state == 0:  # 如果未选中
                            all_checkbox.click_input()
                            time.sleep(0.6)
                            new_done = 1
                            print("✅ '全部' 已勾选")
                        else:
                            print("ℹ️ '全部' 已勾选，跳过")
                            new_done = 1
                    else:
                        all_checkbox.toggle()
                        time.sleep(0.6)
                        new_done = 1
                except Exception as e:
                    print(f"⚠️ 全选失败: {e}")
            else:
                print("❌ 未找到 '全部' CheckBox，无法执行全选")
        else:
            # 勾选新目标店铺
            print("\n🆕 正在勾选新目标店铺...")
            for name in shop_list_names:
                new_done += self._toggle_shop_by_rule(popup, name, popup_shop_name_map)

            print(f"🎉 新目标勾选: {new_done}/{len(shop_list_names)}")

        # === 步骤9: 关闭下拉框 ===
        print("\n关闭下拉框：点击主窗口空白区关闭下拉...")
        mouse.click(coords=(win_rect.left + 20, win_rect.top + 80))
        time.sleep(0.4)
        print("✅ 下拉框已关闭")
        
        return new_done > 0

