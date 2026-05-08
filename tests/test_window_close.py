#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试类目修改窗口关闭功能的脚本
"""

import sys
import os
import time

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from batch_import_automation.workflow import BatchImportAutomation
from category_automation.workflow import CategoryAutomation

def test_window_close():
    """测试类目修改窗口关闭功能"""
    print("🔍 测试类目修改窗口关闭功能...")
    
    try:
        # 创建类目自动化实例
        category_automation = CategoryAutomation()
        print("✅ 成功创建CategoryAutomation实例")
        
        # 查找类目修改窗口
        if not category_automation.find_window():
            print("❌ 未找到类目修改窗口，请先手动打开类目修改窗口")
            return False
        
        print("✅ 找到类目修改窗口")
        
        # 等待一段时间，确保窗口完全加载
        print("⏳ 等待窗口完全加载...")
        time.sleep(2)
        
        # 测试关闭窗口
        print("\n🔘 开始测试关闭窗口...")
        start_time = time.time()
        
        result = category_automation.close_window()
        
        end_time = time.time()
        elapsed = end_time - start_time
        
        if result:
            print(f"✅ 窗口关闭成功，耗时: {elapsed:.2f}秒")
            return True
        else:
            print(f"❌ 窗口关闭失败，耗时: {elapsed:.2f}秒")
            return False
            
    except Exception as e:
        print(f"❌ 测试过程中出错: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_with_batch_automation():
    """通过批量导入自动化测试"""
    print("\n🔍 通过批量导入自动化测试类目修改窗口关闭功能...")
    
    try:
        # 创建批量导入自动化实例
        automation = BatchImportAutomation()
        print("✅ 成功创建BatchImportAutomation实例")
        
        # 查找批量导入窗口
        if not automation.find_window():
            print("❌ 未找到批量导入窗口，请确保批量导入应用程序正在运行")
            return False
        
        # 加载配置
        all_configs = automation.config_manager.get_all_import_configs()
        if not all_configs:
            print("❌ 未找到导入配置，请检查Excel文件")
            return False
        
        # 查找有类目配置的数据
        config_with_category = None
        for config in all_configs:
            category_config = config.get("类目", {})
            if any(category_config.values()):
                config_with_category = config
                break
        
        if not config_with_category:
            print("❌ 未找到有类目配置的数据")
            return False
        
        print(f"📋 找到有类目配置的数据")
        print(f"   站点: {config_with_category.get('站点', '')}")
        print(f"   采集词: {config_with_category.get('采集词', '')}")
        
        # 获取索引
        index = all_configs.index(config_with_category)
        print(f"   数据索引: {index}")
        
        # 只测试类目修改窗口的关闭功能
        print("\n🔘 开始测试类目修改窗口的关闭功能...")
        
        # 点击添加类目按钮
        if not automation.category_automation.click_category_modify_button():
            print("❌ 点击添加类目按钮失败")
            return False
        
        print("✅ 成功点击添加类目按钮")
        
        # 查找类目修改窗口
        if not automation.category_automation.find_window():
            print("❌ 未找到类目修改窗口")
            return False
        
        print("✅ 找到类目修改窗口")
        
        # 等待一段时间，确保窗口完全加载
        print("⏳ 等待窗口完全加载...")
        time.sleep(2)
        
        # 测试关闭窗口
        print("\n🔘 开始测试关闭窗口...")
        start_time = time.time()
        
        result = automation.category_automation.close_window()
        
        end_time = time.time()
        elapsed = end_time - start_time
        
        if result:
            print(f"✅ 窗口关闭成功，耗时: {elapsed:.2f}秒")
            return True
        else:
            print(f"❌ 窗口关闭失败，耗时: {elapsed:.2f}秒")
            return False
            
    except Exception as e:
        print(f"❌ 测试过程中出错: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("类目修改窗口关闭功能测试")
    print("=" * 60)
    
    # 测试选项
    print("\n请选择测试方式:")
    print("1. 直接测试关闭已打开的类目修改窗口")
    print("2. 通过批量导入自动化流程测试")
    
    choice = input("\n请输入选择 (1 或 2): ").strip()
    
    if choice == "1":
        print("\n开始测试方式1: 直接测试关闭已打开的类目修改窗口")
        print("请先手动打开类目修改窗口，然后按任意键继续...")
        input()
        result = test_window_close()
    elif choice == "2":
        print("\n开始测试方式2: 通过批量导入自动化流程测试")
        result = test_with_batch_automation()
    else:
        print("无效的选择")
        result = False
    
    if result:
        print("\n🎉 测试完成！窗口关闭功能正常工作")
    else:
        print("\n⚠️ 测试失败，请检查窗口关闭功能")


