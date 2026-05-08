# 拓客批量导入

Shopee 跨境电商批量上货自动化工具，基于 Windows UI 自动化技术，模拟人工操作「批量导入」桌面客户端，实现商品批量上架的全流程自动化。

## 支持站点

台湾、泰国、马来、新加坡、菲律宾、越南、巴西（7 个站点）

## 功能

- **站点/店铺切换** — 自动切换站点、分组，按规则勾选/取消店铺
- **价格自动计算** — 根据站点 + 重量查表，自动填入基础加价、运费、仓储费（支持 >1kg 分段计算）
- **链接自动加载** — 按站点和采集词从 `links/` 目录匹配链接文件，写入导入框
- **类目修改** — 自动打开类目窗口，选择多级类目，填写属性（品牌/材质/保质期/FDA/TIS/PS-ICC 等），上传认证 PDF
- **模板/尺寸图** — 自动选择导入模板，按需切换尺寸图
- **链接校验容错** — 导入失败时自动检测「多平台链接混用」弹窗，剔除非 TaoWorld 链接后重试
- **失败记录** — 失败的导入自动记录到 `failed_category_links.txt`，方便事后排查
- **多表格批处理** — 支持命令行传入多个 Excel 文件，或自动发现目录下所有 `*批量导入*.xlsx`

## 环境要求

- Windows 10 / 11
- Python 3.10+
- 已打开「批量导入」桌面窗口

## 安装

```bash
pip install -r requirements.txt
```

### 依赖

| 库 | 用途 |
|---|---|
| `uiautomation` | WPF 控件定位与操作 |
| `pywinauto` | 店铺下拉弹窗处理 |
| `openpyxl` | Excel 配置读写 |
| `keyboard` | F5 快捷键监听（可选） |

## 快速开始

### 1. 准备数据

在 `拓客批量导入.xlsx` 中填写商品信息（支持多行）：

| 站点 | 采集词 | 分组 | 店铺 | 重量 | 一级 | 二级 | 三级 | 四级 | 尺寸图 | PDF文件 | ... |
|------|--------|------|------|------|------|------|------|------|--------|---------|-----|
| 台湾 | 棒球领衬衫男 | 男装 | 全部 | 0.5 | 女装 | 上衣 | 背心 | | 尺寸图1 | 化妆品FDA.pdf | ... |

### 2. 准备链接

在 `links/站点名/` 目录下放置链接文件，文件名格式：`采集词_时间戳.txt`，内容为 TaoWorld 链接，每行一个。

```
links/
├── 台湾/
│   ├── 棒球领衬衫男_20250317150748.txt
│   └── ...
├── 泰国/
│   └── ...
└── ...
```

### 3. 配置价格

`价格.json` 已内置 7 个站点的价格公式：

```json
{
    "价格": {
        "台湾": {
            "0.5": "180%+15+10"
        }
    }
}
```

格式：`基础加价%+运费+仓储费`

### 4. 运行

```bash
# 默认处理当前目录下所有 *批量导入*.xlsx
python run_complete.py

# 指定表格
python run_complete.py 表1.xlsx 表2.xlsx
```

程序启动后按 **F5** 开始执行，运行期间不要操作鼠标键盘。

## 项目结构

```
├── run_complete.py              # 主入口
├── batch_import_automation/     # 批量导入核心模块
│   ├── workflow.py              #   主流程编排
│   ├── config.py                #   配置管理（Excel + 价格 JSON）
│   ├── shop_handler.py          #   店铺选择
│   ├── field_handler.py         #   字段填写
│   ├── template_handler.py      #   模板/尺寸图/上传设置
│   ├── file_handler.py          #   链接文件读写
│   └── category_handler.py      #   类目按钮操作
├── category_automation/         # 类目修改子系统
│   ├── workflow.py              #   类目修改主流程
│   ├── dropdown_handler.py      #   下拉框滚动搜索
│   ├── attribute_matcher.py     #   属性标签匹配
│   └── dialog_handler.py        #   弹窗检测处理
├── tools/                       # 辅助工具
│   └── create_cascading_dropdown.py  # 生成 Excel 级联下拉菜单
├── links/                       # 链接数据（按站点分目录）
├── 价格.json                     # 多站点价格公式
└── 拓客批量导入.xlsx              # 商品数据表格
```

## 打包

```bash
pip install pyinstaller
pyinstaller 批量导入.spec
```

输出到 `dist/批量导入.exe`。

## License

MIT
