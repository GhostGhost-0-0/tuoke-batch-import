# 类目修改自动化脚本

## 简介

这是一个用于自动化操作"类目修改"窗口的Python脚本，支持：
- 自动选择多级类目（一级、二级、三级、四级）
- 自动加载类目属性
- 自动设置属性值（支持下拉框和输入框）
- 自动处理弹窗对话框
- 自动点击确认按钮

## 项目结构

```
批量导入/
├── category_automation/          # 主模块目录
│   ├── __init__.py               # 模块初始化
│   ├── config.py                 # 配置管理模块
│   ├── ui_utils.py               # UI工具模块（鼠标、键盘、延迟等）
│   ├── dropdown_handler.py      # 下拉框处理模块
│   ├── dialog_handler.py         # 弹窗处理模块
│   ├── attribute_matcher.py      # 属性匹配模块
│   └── workflow.py                # 主流程模块
├── run.py                        # 入口文件
├── test.py                       # 原始脚本（保留作为参考）
├── 属性配置.json                  # 配置文件
├── requirements.txt              # 依赖包列表
└── README.md                     # 本文件
```

## 安装依赖

```bash
pip install -r requirements.txt
```

## 配置说明

编辑 `属性配置.json` 文件来配置类目和属性：

```json
{
  "类目": {
    "一级": "美妆保健(美妆保健)",
    "二级": "手足保养与美甲(手足保养与美甲)",
    "三级": "手部保养(手部保养)",
    "四级": "手膜"
  },
  "属性": {
    "品牌": "chloe",
    "保质期": "24个月"
  }
}
```

### 配置项说明

- **类目配置**：
  - `一级`、`二级`、`三级`：必需项，填写完整的类目名称（包含括号内容）
  - `四级`：可选项，如果留空则跳过四级类目选择

- **属性配置**：
  - 键名：属性名称（如"品牌"、"保质期"等）
  - 值：要设置的属性值
  - 支持下拉框和输入框两种类型的属性

## 使用方法

### 方法1：使用入口文件（推荐）

```bash
python run.py
```

### 方法2：作为模块使用

```python
from category_automation import CategoryAutomation

# 创建自动化实例
automation = CategoryAutomation(config_file="属性配置.json")

# 执行自动化流程
automation.run()
```

### 方法3：分步执行

```python
from category_automation import CategoryAutomation

automation = CategoryAutomation(config_file="属性配置.json")

# 1. 查找窗口
if automation.find_window():
    # 2. 最大化窗口
    automation.maximize_window()
    
    # 3. 加载配置
    config = automation.config_manager.load()
    
    # 4. 选择类目
    automation.select_categories(config.get("类目", {}))
    
    # 5. 加载属性
    automation.load_attributes()
    
    # 6. 处理属性
    results = automation.process_attributes(config.get("属性", {}))
    
    # 7. 点击确认
    automation.click_confirm()
    
    # 8. 保存结果
    automation.save_results(results)
```

## 模块说明

### ConfigManager（配置管理）

负责加载和管理配置文件：

```python
from category_automation import ConfigManager

config_manager = ConfigManager("属性配置.json")
config = config_manager.load()
category_config = config_manager.get_category_config()
attribute_config = config_manager.get_attribute_config()
```

### CategoryAutomation（主流程）

主自动化类，提供完整的自动化流程：

```python
from category_automation import CategoryAutomation

automation = CategoryAutomation()
automation.run()  # 执行完整流程
```

## 功能特性

1. **智能类目选择**：使用拼音输入快速选择类目
2. **下拉框滚动查找**：自动滚动下拉框查找目标选项
3. **属性自动匹配**：智能匹配属性名称和控件
4. **弹窗自动处理**：自动检测并处理弹窗对话框
5. **安全鼠标操作**：确保鼠标操作在窗口范围内
6. **详细日志输出**：提供详细的执行日志

## 注意事项

1. 运行前请确保"类目修改"窗口已打开
2. 配置文件使用UTF-8编码
3. 脚本会自动最大化窗口以确保操作准确性
4. 属性名称支持模糊匹配（部分匹配）

## 故障排除

### 问题：找不到窗口

**解决方案**：确保"类目修改"窗口已打开且可见

### 问题：找不到属性字段

**解决方案**：检查属性名称是否正确，脚本支持模糊匹配

### 问题：下拉框选项找不到

**解决方案**：检查配置值是否完全匹配下拉框中的选项文本（不区分大小写）

## 更新日志

### v1.0.0
- 初始版本
- 模块化重构
- 保留所有原有功能
- 优化代码结构

## 许可证

本项目仅供学习和个人使用。

