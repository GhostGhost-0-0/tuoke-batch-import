# 打包说明

## 方法一：使用批处理脚本打包（推荐）

### Windows 系统

1. **双击运行 `build.bat`**
   - 脚本会自动安装 PyInstaller（如果未安装）
   - 自动安装依赖项
   - 自动打包成可执行文件

2. **打包完成后**
   - 可执行文件位置：`dist\提取类目文本.exe`
   - 将 `提取类目文本.exe` 复制到目标电脑即可使用

### 手动打包步骤

如果批处理脚本无法运行，可以手动执行：

```bash
# 1. 安装 PyInstaller
pip install pyinstaller

# 2. 安装依赖项
pip install -r requirements.txt

# 3. 打包
pyinstaller --onefile --name "提取类目文本" --console extract_categories.py
```

## 方法二：直接复制 Python 脚本

如果目标电脑已安装 Python，可以直接复制脚本文件：

1. **复制文件**
   - `extract_categories.py`
   - `requirements.txt`

2. **在目标电脑上安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

3. **运行脚本**
   ```bash
   python extract_categories.py
   ```

## 依赖项说明

脚本需要以下 Python 包：

- `uiautomation` - Windows UI 自动化库
- `pywinauto` - Windows 应用程序自动化库

标准库（无需安装）：
- `time`
- `json`
- `ctypes`
- `typing`

## 使用要求

1. **操作系统**：Windows 7 及以上版本
2. **Python 版本**：Python 3.6 及以上（如果使用 Python 脚本）
3. **目标应用程序**：需要先打开"类目修改"窗口

## 打包后的文件结构

```
dist/
  └── 提取类目文本.exe    # 可执行文件（约 20-30MB）
```

## 注意事项

1. **首次运行**：打包后的 exe 文件可能被杀毒软件拦截，需要添加信任
2. **文件位置**：exe 文件会在同目录下创建 `文本.json` 文件
3. **窗口名称**：确保"类目修改"窗口已打开且可见
4. **权限**：某些情况下可能需要管理员权限运行

## 故障排除

### 问题1：打包失败
- 确保已安装 Python 3.6+
- 确保 pip 可以正常使用
- 尝试手动安装：`pip install pyinstaller uiautomation pywinauto`

### 问题2：exe 文件无法运行
- 检查是否被杀毒软件拦截
- 尝试以管理员身份运行
- 检查目标电脑是否有必要的运行库（通常 Windows 10+ 自带）

### 问题3：找不到窗口
- 确保"类目修改"窗口已打开
- 确保窗口标题完全匹配
- 尝试先手动激活窗口

## 联系支持

如有问题，请检查：
1. Python 版本和依赖项是否正确安装
2. 目标应用程序是否正常运行
3. 是否有足够的系统权限















