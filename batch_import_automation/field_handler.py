"""
字段设置处理模块
"""

import uiautomation as auto
import time
from .ui_utils import set_control_value, human_delay


def _link_value_console_repr(value: str) -> str:
    """控制台展示用：多行链接只打印首条，其余用省略说明。"""
    if value is None or not str(value).strip():
        return "(空)"
    lines = [
        ln.strip()
        for ln in str(value).replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if ln.strip()
    ]
    if not lines:
        return "(空)"
    if len(lines) == 1:
        return lines[0]
    return f"{lines[0]} ... (共 {len(lines)} 行)"


class FieldHandler:
    """字段设置处理类"""
    
    # 字段名到AutomationId的映射
    FIELD_AUTOMATION_IDS = {
        "关键词": "HotSearchskinTextBox",
        "热搜词": "HotSearchskinTextBox",  # 支持两种名称
        "基础加价": "JiChuskinTextBox",
        "运费": None,  # 需要通过标签查找
        "重量": "MinWeightskinTextBox",  # 重量字段
        "链接": "skinTextBox1",  # 链接输入框
    }
    
    def __init__(self, window=None):
        """
        初始化字段处理器
        :param window: 窗口控件对象
        """
        self.window = window
    
    def find_control_by_name(self, name: str, control_type=None):
        """
        通过名称查找控件
        :param name: 控件名称（支持部分匹配）
        :param control_type: 控件类型（可选）
        :return: 控件对象或None
        """
        if not self.window:
            return None
        
        try:
            # 获取所有后代控件
            def search_controls(ctrl, depth=0):
                if depth > 10:
                    return None
                try:
                    ctrl_name = ctrl.Name or ""
                    if name in ctrl_name or ctrl_name in name:
                        if control_type is None or ctrl.ControlType == control_type:
                            return ctrl
                except:
                    pass
                
                try:
                    child = ctrl.GetFirstChildControl()
                    while child:
                        result = search_controls(child, depth + 1)
                        if result:
                            return result
                        child = child.GetNextSiblingControl()
                except:
                    pass
                return None
            
            return search_controls(self.window)
        except:
            return None
    
    def find_input_by_label(self, label_text: str):
        """
        通过标签文本查找对应的输入控件
        :param label_text: 标签文本（支持部分匹配）
        :return: (标签控件, 输入控件) 或 (None, None)
        """
        if not self.window:
            return None, None
        
        try:
            def get_all_controls(ctrl, depth=0, max_depth=10):
                if depth > max_depth:
                    return []
                controls = [ctrl]
                try:
                    child = ctrl.GetFirstChildControl()
                    while child:
                        controls.extend(get_all_controls(child, depth + 1, max_depth))
                        child = child.GetNextSiblingControl()
                except:
                    pass
                return controls
            
            all_controls = get_all_controls(self.window)
            
            # 查找标签
            labels = []
            inputs = []
            
            for ctrl in all_controls:
                try:
                    name = (ctrl.Name or "").strip()
                    if (ctrl.ControlType == auto.ControlType.TextControl and 
                        label_text in name and len(name) <= 20):
                        labels.append(ctrl)
                    elif (ctrl.ControlType in [auto.ControlType.EditControl, 
                                               auto.ControlType.ComboBoxControl] and
                          not ctrl.IsOffscreen and ctrl.IsEnabled):
                        inputs.append(ctrl)
                except:
                    continue
            
            # 匹配标签和输入控件
            for label in labels:
                label_rect = label.BoundingRectangle
                best_input = None
                min_distance = float('inf')
                
                for inp in inputs:
                    try:
                        inp_rect = inp.BoundingRectangle
                        if inp_rect.width() <= 0 or inp_rect.height() <= 0:
                            continue
                        
                        # 检查是否在同一水平线上
                        if (inp_rect.top - 10 < label_rect.bottom and 
                            inp_rect.bottom + 10 > label_rect.top):
                            # 计算距离
                            dx = inp_rect.left - label_rect.right
                            if 0 < dx < 500:
                                distance = abs((inp_rect.top + inp_rect.bottom) / 2 - 
                                              (label_rect.top + label_rect.bottom) / 2)
                                if distance < min_distance:
                                    min_distance = distance
                                    best_input = inp
                    except:
                        continue
                
                if best_input:
                    return label, best_input
            
            return None, None
        except Exception as e:
            print(f"  ⚠️ 查找控件失败: {e}")
            return None, None
    
    def set_field_value(self, field_name: str, value: str) -> bool:
        """
        设置字段值
        :param field_name: 字段名称（如"采集词"、"关键词"等）
        :param value: 要设置的值
        :return: 是否成功设置
        """
        log_value = _link_value_console_repr(value) if field_name == "链接" else value
        print(f"\n🔄 设置字段: '{field_name}' = '{log_value}'")
        
        # 特殊处理：基础加价需要去掉百分号
        if field_name == "基础加价" and value.endswith("%"):
            value = value.rstrip("%")
            print(f"  ℹ️ 基础加价去掉百分号: {value}")
            log_value = value
        
        # 特殊处理：重量字段需要同时填写最小重量和最大重量两个框
        if field_name == "重量":
            return self._set_weight_value(value)
        
        input_ctrl = None
        
        # 方法1: 优先使用AutomationId查找（最可靠）
        automation_id = self.FIELD_AUTOMATION_IDS.get(field_name)
        if automation_id:
            try:
                input_ctrl = self.window.EditControl(AutomationId=automation_id)
                if input_ctrl.Exists(maxSearchSeconds=2):
                    print(f"  ✓ 通过AutomationId找到字段 '{field_name}' (ID: {automation_id})")
                else:
                    input_ctrl = None
            except:
                input_ctrl = None
        
        # 方法2: 如果AutomationId未找到，通过标签查找
        if not input_ctrl:
            label, input_ctrl = self.find_input_by_label(field_name)
            if input_ctrl:
                print(f"  ✓ 通过标签找到字段 '{field_name}'")
        
        if input_ctrl:
            # 设置值
            if input_ctrl.ControlType == auto.ControlType.EditControl:
                if set_control_value(input_ctrl, value):
                    print(f"  ✅ 已设置 {field_name} 为: {log_value}")
                    return True
            elif input_ctrl.ControlType == auto.ControlType.ComboBoxControl:
                # 下拉框需要特殊处理
                print(f"  ⚠️ 下拉框类型，需要特殊处理")
                return False
            else:
                print(f"  ⚠️ 不支持的控件类型: {input_ctrl.ControlType}")
        else:
            print(f"  ⚠️ 未找到字段 '{field_name}'")
            if automation_id:
                print(f"    尝试的AutomationId: {automation_id}")
        
        return False
    
    def _set_weight_value(self, value: str) -> bool:
        """
        设置重量值（同时填写最小重量和最大重量两个框）
        :param value: 重量值
        :return: 是否成功设置
        """
        if not self.window:
            print("  ⚠️ 窗口未初始化，无法设置重量")
            return False
        
        success_count = 0
        
        # 设置最小重量
        try:
            min_weight_ctrl = self.window.EditControl(AutomationId="MinWeightskinTextBox")
            if min_weight_ctrl.Exists(maxSearchSeconds=2):
                print(f"  ✓ 找到最小重量框 (MinWeightskinTextBox)")
                if set_control_value(min_weight_ctrl, value):
                    print(f"  ✅ 已设置最小重量为: {value}")
                    success_count += 1
                else:
                    print(f"  ⚠️ 设置最小重量失败")
            else:
                print(f"  ⚠️ 未找到最小重量框")
        except Exception as e:
            print(f"  ⚠️ 查找最小重量框时出错: {e}")
        
        # 在两个重量框之间添加短暂延迟
        human_delay(0.2, 0.3)
        
        # 设置最大重量
        try:
            max_weight_ctrl = self.window.EditControl(AutomationId="MaxWeightskinTextBox")
            if max_weight_ctrl.Exists(maxSearchSeconds=2):
                print(f"  ✓ 找到最大重量框 (MaxWeightskinTextBox)")
                if set_control_value(max_weight_ctrl, value):
                    print(f"  ✅ 已设置最大重量为: {value}")
                    success_count += 1
                else:
                    print(f"  ⚠️ 设置最大重量失败")
            else:
                print(f"  ⚠️ 未找到最大重量框")
        except Exception as e:
            print(f"  ⚠️ 查找最大重量框时出错: {e}")
        
        # 如果两个框都设置成功，返回True
        if success_count == 2:
            print(f"  ✅ 重量字段设置完成（最小重量和最大重量都已设置）")
            return True
        elif success_count == 1:
            print(f"  ⚠️ 重量字段部分设置成功（仅设置了{success_count}个框）")
            return True  # 至少设置了一个，也算成功
        else:
            print(f"  ❌ 重量字段设置失败（两个框都未设置成功）")
            return False

