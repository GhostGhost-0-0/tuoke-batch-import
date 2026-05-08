"""
批量导入自动化测试入口文件
只执行批量导入流程，不执行类目修改
"""

from batch_import_automation import BatchImportAutomation


def main():
    """主函数：只执行批量导入流程"""
    print("=" * 60)
    print("🚀 开始批量导入自动化测试")
    print("=" * 60)
    
    config_file = "属性配置.json"
    
    # 创建批量导入自动化实例
    batch_automation = BatchImportAutomation(config_file=config_file)
    
    # 执行批量导入流程
    success = batch_automation.run()
    
    if success:
        print("\n✅ 批量导入流程执行成功！")
    else:
        print("\n❌ 批量导入流程执行失败，请检查日志文件 batch_import.log")
    
    return success


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)










