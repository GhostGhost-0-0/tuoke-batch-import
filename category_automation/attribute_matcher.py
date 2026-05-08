"""
属性匹配器模块
用于匹配属性标签和输入控件
"""

import uiautomation as auto
from .ui_utils import get_all_descendants


def pair_label_with_input(window):
    """
    匹配标签和输入控件
    :param window: 窗口对象
    :return: 匹配的(标签, 输入控件)对列表
    """
    labels = []
    inputs = []
    all_controls = get_all_descendants(window)
    keywords = ["材质", "尺寸（长 x 宽 x 高）", "适用年龄", "FDA化妆品注册编号", "FDA", "TIS标准编号", "FDA医疗器械注册号", "PS/ICC编号"]
    
    for ctrl in all_controls:
        name = (ctrl.Name or "").strip()
        if ctrl.ControlType == auto.ControlType.TextControl and any(k in name for k in keywords):
            # FDA 相关标签可能较长（如英文全称），放宽长度限制
            max_len = 80 if "FDA" in name else 20
            if 1 <= len(name) <= max_len:
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


def find_new_controls(window, old_snapshot):
    """
    查找新增控件（通过对比快照）
    :param window: 窗口对象
    :param old_snapshot: 旧快照
    :return: 新增控件列表
    """
    from .ui_utils import take_snapshot
    
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




