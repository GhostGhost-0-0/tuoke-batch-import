#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试重量字段处理功能
"""

import os
import sys
import json

# 添加batch_import_automation目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'batch_import_automation'))

from batch_import_automation.config import BatchImportConfigManager
from batch_import_automation.field_handler import FieldHandler

def test_weight_field():
    """测试重量字段的处理"""
    print("测试重量字段处理功能")
    print("="*50)
    
    # 1. 测试配置文件是否包含重量字段
    config_manager = BatchImportConfigManager()
    excel_data = config_manager.load_excel_data()
    
    if excel_data and '重量' in excel_data:
        weight_value = excel_data.get('重量')
        print(f"✓ 配置文件包含重量字段，值为: {weight_value}")
    else:
        print("✗ 配置文件未包含重量字段或值为空")
        return False
    
    # 2. 测试FieldHandler是否能识别重量字段
    field_handler = FieldHandler()
    automation_id = field_handler.FIELD_AUTOMATION_IDS.get('重量')
    if automation_id:
        print(f"✓ FieldHandler包含重量字段的AutomationId: {automation_id}")
    else:
        print("✗ FieldHandler未包含重量字段的AutomationId")
        return False
    
    print("="*50)
    print("所有测试通过！重量字段处理功能已正确配置。")
    return True

if __name__ == "__main__":
    try:
        test_weight_field()
    except Exception as e:
        print(f"测试过程中出错: {e}")
        import traceback
        traceback.print_exc()
