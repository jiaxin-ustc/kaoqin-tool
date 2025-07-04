# 🕒 考勤异常检测系统

## 📋 项目简介
本项目基于 **Gradio** 框架与 **Pandas** 库，实现了一个简洁高效的考勤异常检测 C/S 系统。通过 PyInstaller 打包为 exe，用户可在 Windows 环境下直接运行，无需安装 Python。

✨ 管理员可通过网页上传员工打卡数据，系统自动分析，识别迟到、早退、缺卡等常见异常，并以红色高亮展示，支持一键下载结果和异常统计报告。

---

## 🚀 主要功能
- **结果写回原表格**：分析结果直接写回原始表格，便于后续处理。
- **异常红色标记**：异常数据以红色背景突出显示。
- **智能异常检测**：
  - 自动识别迟到、早退、缺卡等情况
  - 支持多种打卡时间格式解析
  - 智能分配时间点到对应时间段
- **精确标记定位**：
  - 准确定位异常单元格
  - 保持原有数据格式和结构
  - 添加详细异常说明
- **安全文件处理**：
  - 自动生成新文件，不覆盖原始数据
  - 完善的错误处理与日志输出
  - 支持自定义输出路径

---

## ⚠️ 异常类型说明

1. ⏰ **上午迟到**：上班时间晚于 8:00
2. 🏃‍♂️ **上午早退**：下班时间早于 12:00
3. ⏰ **下午迟到**：下午上班晚于 13:00
4. 🏃‍♂️ **下午早退**：下午下班早于 15:30
5. ❌ **缺卡**：对应时间段无打卡记录

---

## 🛠️ 使用方法
1. ⬇️ 下载 exe 文件，双击运行。
2. 🌐 浏览器访问 [http://localhost:7860/](http://localhost:7860/)，上传考勤数据文件。
3. 🖱️ 点击"分析数据"按钮，系统自动处理并展示结果。

---

## �� 本地运行与打包

### 本地运行
```bash
# 进入项目目录
cd path/to/your/project
# 创建虚拟环境，安装依赖
conda create -n myenv python=3.9
pip install -r requirements.txt
# 运行程序
python main.py
```

### 项目打包
```bash
# 安装打包工具
pip install pyinstaller
# 生成 spec 文件
pyi-makespec --onefile --collect-all gradio --collect-all gradio_client main.py
# 修改 spec 文件，添加 gradio 的编译
a = Analysis(
    ['app.spec'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    module_collection_mode={ 'gradio': 'py',},
    optimize=0,
)
# 继续打包
pyinstaller main.spec
```

> ⚡ **注意**：如需自定义打包细节，请参考 `main.spec` 文件中的 `Analysis` 配置。

---

## 📦 输出文件说明
- `kaohe_result.xlsx`：详细异常记录报告
- `原文件名_marked.xlsx`：标记异常的原表格文件

---

## 🙋 常见问题
- 如遇依赖缺失或打包异常，请检查 Python 版本与依赖包是否齐全。
- 建议使用最新版 PyInstaller 及 Gradio。

---

## ©️ 版权声明

版权所有 © 2025 Jiaxin Jiang. 保留所有权利。
