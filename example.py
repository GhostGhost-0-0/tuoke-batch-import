"""
使用示例文件
展示如何使用类目修改自动化模块
"""

from category_automation import CategoryAutomation, ConfigManager


def example_simple():
    """简单使用示例"""
    print("=" * 50)
    print("示例1：简单使用")
    print("=" * 50)
    
    # 创建自动化实例并执行
    automation = CategoryAutomation(config_file="属性配置.json")
    automation.run()


def example_step_by_step():
    """分步执行示例"""
    print("\n" + "=" * 50)
    print("示例2：分步执行")
    print("=" * 50)
    
    automation = CategoryAutomation(config_file="属性配置.json")
    
    # 1. 查找窗口
    if not automation.find_window():
        print("未找到窗口，退出")
        return
    
    # 2. 最大化窗口
    automation.maximize_window()
    
    # 3. 加载配置
    config = automation.config_manager.load()
    category_config = config.get("类目", {})
    attribute_config = config.get("属性", {})
    
    # 4. 选择类目
    if not automation.select_categories(category_config):
        print("类目选择失败，退出")
        return
    
    # 5. 加载属性
    if not automation.load_attributes():
        print("加载属性失败，退出")
        return
    
    # 6. 处理属性
    results = automation.process_attributes(attribute_config)
    
    # 7. 点击确认
    if attribute_config:
        automation.click_confirm()
    
    # 8. 保存结果
    automation.save_results(results)


def example_custom_config():
    """自定义配置示例"""
    print("\n" + "=" * 50)
    print("示例3：自定义配置")
    print("=" * 50)
    
    # 使用自定义配置文件
    automation = CategoryAutomation(config_file="自定义配置.json")
    
    # 或者动态修改配置
    config_manager = ConfigManager("属性配置.json")
    config = config_manager.load()
    
    # 修改配置
    config["属性"]["品牌"] = "新品牌"
    config["属性"]["保质期"] = "36个月"
    
    # 执行
    automation.config_manager._config = config
    automation.run()


if __name__ == "__main__":
    # 运行简单示例
    example_simple()
    
    # 取消注释以运行其他示例
    # example_step_by_step()
    # example_custom_config()

