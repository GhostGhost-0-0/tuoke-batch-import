"""
类目修改自动化脚本入口文件
"""

import sys
from category_automation import CategoryAutomation


def main():
    """主函数"""
    print("提示: 如果需要控制启动，请使用 run_complete.py")

    # 创建自动化实例
    automation = CategoryAutomation(config_file="属性配置.json")

    # 执行自动化流程
    try:
        success = automation.run()
    except KeyboardInterrupt:
        print("\n\n⚠️ 收到中断信号")
        success = False
    except Exception as e:
        print(f"\n\n❌ 执行过程中发生错误: {e}")
        success = False

    if success:
        print("\n✅ 自动化流程执行成功！")
    else:
        print("\n❌ 自动化流程执行失败！")

    return success


if __name__ == "__main__":
    success = main()

    # 等待用户输入，防止控制台自动关闭
    print("\n" + "=" * 60)
    print("程序执行完成，按 Enter 键退出...")
    try:
        input()
    except (EOFError, KeyboardInterrupt):
        pass


