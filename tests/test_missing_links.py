#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试未找到链接文件时记录并跳过的功能
"""

import sys
import os

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from batch_import_automation.workflow import BatchImportAutomation


def _line_site_cw(parts):
    """与 workflow 失败行格式一致: 时间|站点|采集词|原因 或 旧格式 站点|采集词。"""
    if len(parts) >= 4:
        return parts[1], parts[2]
    if len(parts) >= 2:
        return parts[0], parts[1]
    return None, None


def test_missing_links():
    """测试未找到链接文件的处理"""
    print("🔍 测试未找到链接文件的处理...")
    
    try:
        # 创建自动化实例
        automation = BatchImportAutomation()
        print("✅ 成功创建BatchImportAutomation实例")
        
        # 查找窗口
        if not automation.find_window():
            print("❌ 未找到批量导入窗口，请确保批量导入应用程序正在运行")
            return False
        
        # 加载配置
        all_configs = automation.config_manager.get_all_import_configs()
        if not all_configs:
            print("❌ 未找到导入配置，请检查Excel文件")
            return False
        
        # 查找有链接需求的配置
        config_with_links = None
        for config in all_configs:
            site = config.get("站点", "")
            collection_word = config.get("采集词", "")
            if site and collection_word:
                config_with_links = config
                break
        
        if not config_with_links:
            print("❌ 未找到有链接需求的配置")
            return False
        
        print(f"📋 找到有链接需求的配置:")
        print(f"   站点: {config_with_links.get('站点', '')}")
        print(f"   采集词: {config_with_links.get('采集词', '')}")
        
        # 获取索引
        index = all_configs.index(config_with_links)
        print(f"   数据索引: {index}")
        
        # 临时修改采集词为一个不存在的值，模拟链接文件不存在的情况
        original_collection_word = config_with_links.get('采集词', '')
        config_with_links['采集词'] = '不存在的采集词_测试用'
        
        print(f"\n🔘 开始测试未找到链接文件的情况...")
        print(f"   模拟使用不存在的采集词: {config_with_links['采集词']}")
        
        # 测试处理单条数据配置
        result = automation.process_single_row_config(config_with_links, index + 1)
        
        # 检查是否创建了失败链接文件
        failed_links_file = "failed_category_links.txt"
        if os.path.exists(failed_links_file):
            print(f"\n✅ 已创建失败链接文件: {failed_links_file}")
            
            # 读取并显示失败链接内容
            try:
                with open(failed_links_file, "r", encoding="utf-8") as f:
                    lines = [line.strip() for line in f.readlines() if line.strip()]
                
                print(f"\n📄 失败记录 (共 {len(lines)} 条):")
                for i, line in enumerate(lines[-5:], 1):  # 显示最后5条
                    parts = line.split("|")
                    site_p, cw_p = _line_site_cw(parts)
                    if site_p and cw_p:
                        if len(parts) >= 4:
                            reason = "|".join(parts[3:])
                            rshow = reason[:60] + "..." if len(reason) > 60 else reason
                            print(f"   {i}. 站点: {site_p}, 采集词: {cw_p}, 原因: {rshow}")
                        else:
                            print(f"   {i}. 站点: {site_p}, 采集词: {cw_p} (旧格式)")
                    else:
                        print(f"   {i}. {line}")

                # 检查是否包含我们刚刚测试的记录
                test_record_found = False
                for line in lines:
                    parts = line.split("|")
                    site_p, cw_p = _line_site_cw(parts)
                    if (
                        site_p == config_with_links["站点"]
                        and cw_p == config_with_links["采集词"]
                    ):
                        test_record_found = True
                        break
                
                if test_record_found:
                    print("\n✅ 测试记录已正确保存到失败链接文件")
                else:
                    print("\n⚠️ 测试记录未在失败链接文件中找到")
                    
            except Exception as e:
                print(f"  ⚠️ 读取失败链接文件时出错: {e}")
        else:
            print(f"\n⚠️ 未创建失败链接文件: {failed_links_file}")
        
        # 恢复原始采集词
        config_with_links['采集词'] = original_collection_word
        
        # 返回测试结果
        if result:
            print("\n❌ 测试失败：process_single_row_config 应该返回 False")
            return False
        else:
            print("\n✅ 测试成功：process_single_row_config 正确返回 False")
            return True
            
    except Exception as e:
        print(f"❌ 测试过程中出错: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("未找到链接文件测试")
    print("=" * 60)
    
    result = test_missing_links()
    
    if result:
        print("\n🎉 测试完成！未找到链接文件的处理功能正常工作")
    else:
        print("\n⚠️ 测试过程中发现问题，请检查")
