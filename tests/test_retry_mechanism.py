#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试类目修改重试机制的简单脚本
"""

import sys
import os

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from batch_import_automation.workflow import BatchImportAutomation

def test_retry_mechanism():
    """测试类目修改的重试机制"""
    print("🔍 测试类目修改重试机制...")
    
    # 创建自动化实例
    try:
        automation = BatchImportAutomation()
        print("✅ 成功创建BatchImportAutomation实例")
    except Exception as e:
        print(f"❌ 创建实例失败: {e}")
        return False
    
    # 查找窗口
    if not automation.find_window():
        print("❌ 未找到批量导入窗口，请确保批量导入应用程序正在运行")
        return False
    
    # 加载配置
    all_configs = automation.config_manager.get_all_import_configs()
    if not all_configs:
        print("❌ 未找到导入配置，请检查Excel文件")
        return False
    
    print(f"📋 共找到 {len(all_configs)} 条配置")
    
    # 测试第一组数据的类目修改（模拟失败情况）
    first_config = all_configs[0]
    print(f"\n📦 测试处理第一条数据...")
    print(f"   站点: {first_config.get('站点', '')}")
    print(f"   采集词: {first_config.get('采集词', '')}")
    
    # 检查是否有类目配置
    category_config = first_config.get("类目", {})
    if not any(category_config.values()):
        print("⚠️ 第一条数据没有类目配置，请检查Excel文件")
        return False
    
    print(f"   类目配置: {category_config}")
    
    # 测试类目修改
    print("\n🔘 开始测试类目修改处理...")
    success = automation.category_automation.process_category_for_row(0)
    
    if success:
        print("✅ 类目修改处理成功")
        return True
    else:
        print("❌ 类目修改处理失败，检查重试机制是否正常工作")
        print("   检查是否创建了failed_category_links.txt文件")
        return False

if __name__ == "__main__":
    try:
        result = test_retry_mechanism()
        if result:
            print("\n🎉 测试完成！")
        else:
            print("\n⚠️ 测试过程中发现问题，请检查")
    except Exception as e:
        print(f"\n❌ 测试过程中出错: {e}")
        import traceback
        traceback.print_exc()


