"""
完整自动化流程入口文件
执行批量导入流程（包括类目修改）
"""

import time
import sys
from batch_import_automation import BatchImportAutomation

try:
    import keyboard
    HAS_KEYBOARD = True
except ImportError:
    HAS_KEYBOARD = False
    print("⚠️ 未安装 keyboard 库，将跳过快捷键功能")
    print("   安装方法: pip install keyboard")


def wait_for_f5():
    """等待用户按下 F5 键"""
    if not HAS_KEYBOARD:
        print("\n⚠️ keyboard 库不可用，直接开始运行...")
        return

    print("\n" + "=" * 60)
    print("🔴 等待启动")
    print("=" * 60)
    print("\n按 [F5] 键开始运行...")
    print("按 [Ctrl+C] 键退出\n")

    try:
        keyboard.wait('f5')
        print("✅ 已按下 F5，开始运行！\n")
    except KeyboardInterrupt:
        print("\n\n❌ 用户取消，程序退出")
        sys.exit(0)


def main():
    """主函数：执行完整的自动化流程"""
    # 等待用户按下 F5 键
    wait_for_f5()

    print("=" * 60)
    print("🚀 开始完整自动化流程")
    print("=" * 60)
    print("\n批量导入流程（包含类目修改）\n")

    total_start = time.perf_counter()
    config_file = "属性配置.json"
    # 多表格：在下方列表按顺序填写路径；留空则用默认「批量导入.xlsx」
    # 或运行时: python run_complete.py 表1.xlsx 表2.xlsx
    excel_files_override = None  # 例如: [r"g:\data\第一批.xlsx", r"g:\data\第二批.xlsx"]
    if len(sys.argv) > 1:
        excel_files_override = sys.argv[1:]

    # 批量导入流程（现在包含类目修改）
    print("\n" + "=" * 60)
    print("📦 批量导入流程（包含类目修改）")
    print("=" * 60)

    batch_automation = BatchImportAutomation(
        config_file=config_file,
        excel_files=excel_files_override,
        # 未通过命令行指定文件时：当前目录下所有「*批量导入*.xlsx」按文件名顺序依次处理（含 拓客批量导入.xlsx / 拓客批量导入1.xlsx）
        auto_discover_batch_excels=(excel_files_override is None),
    )

    try:
        # run() 方法会自动处理所有配置（包括类目修改）
        batch_success = batch_automation.run()
    except KeyboardInterrupt:
        print("\n\n⚠️ 收到中断信号")
        batch_success = False
    except Exception as e:
        print(f"\n\n❌ 执行过程中发生错误: {e}")
        batch_success = False

    # === 总结 ===
    total_end = time.perf_counter()
    total_time = total_end - total_start

    print("\n" + "=" * 60)
    if batch_success:
        print("🎉 批量导入流程执行成功！")
    else:
        print("❌ 批量导入流程执行失败")
    print("=" * 60)
    print(f"📊 总耗时: {total_time:.2f} 秒")
    print(f"   - 批量导入（含类目修改）: {'✅ 成功' if batch_success else '❌ 失败'}")
    print("=" * 60)

    return batch_success


if __name__ == "__main__":
    success = main()
    if success:
        print("\n✅ 所有流程执行成功！")
    else:
        print("\n❌ 部分流程执行失败，请检查日志")
    
    # 等待用户输入，防止控制台自动关闭
    print("\n" + "=" * 60)
    print("程序执行完成，按 Enter 键退出...")
    try:
        input()
    except (EOFError, KeyboardInterrupt):
        pass

