# 考勤异常检测系统 - 新功能说明

## 新增功能

### 1. 结果写回原表格功能
- **功能描述**: 将考核评估结果直接写回到原始Excel表格中
- **实现方式**: 使用openpyxl库操作Excel文件，保持原有格式不变
- **输出文件**: 自动生成带"_marked"后缀的新文件，避免覆盖原文件

### 2. 异常记录红色标记
- **标记方式**: 存在考勤问题的单元格会被标记为红色背景，白色字体
- **异常信息**: 在原始打卡时间后添加异常详情说明
- **格式示例**: 
  ```
  07:35 12:08
  [异常: 上午迟到 | 下午上班缺卡]
  ```

## 使用方法

### 方法1: 直接调用函数
```python
from kaoqin_hexin_chognxie import analyze_attendance

# 分析考勤数据并写回原表格
result_message, detailed_report, summary, write_back_result = analyze_attendance("your_file.xlsx")
print(write_back_result)
```

### 方法2: 单独使用写回功能
```python
from kaoqin_hexin_chognxie import write_results_back_to_excel
import pandas as pd

# 假设已有处理结果
result_df = pd.read_excel("kaohe_result.xlsx")

# 写回原表格
result = write_results_back_to_excel("original_file.xlsx", result_df, "output_file.xlsx")
print(result)
```

### 方法3: 使用Gradio界面
```python
# 取消注释Gradio界面代码，启动Web界面
# 上传文件后点击"分析数据"按钮即可
```

## 功能特点

### 智能异常检测
- 自动识别迟到、早退、缺卡等异常情况
- 支持多种打卡时间格式解析
- 智能分配时间点到对应时间段

### 精确标记定位
- 准确定位到原始表格中的异常单元格
- 保持原有数据格式和结构
- 添加详细的异常说明信息

### 安全文件处理
- 自动生成新文件，不覆盖原始数据
- 完整的错误处理和日志输出
- 支持自定义输出文件路径

## 异常类型说明

1. **上午迟到**: 上午上班时间晚于8:00
2. **上午早退**: 上午下班时间早于12:00
3. **下午迟到**: 下午上班时间晚于13:00
4. **下午早退**: 下午下班时间早于15:30
5. **缺卡**: 对应时间段没有打卡记录

## 输出文件说明

- `kaohe_result.xlsx`: 详细异常记录报告
- `原文件名_marked.xlsx`: 标记了异常的原表格文件

## 依赖库

```bash
pip install pandas numpy openpyxl gradio
```

## 注意事项

1. 确保原始Excel文件格式正确（第4行为日期行，第5行开始为数据）
2. 员工姓名必须与原始表格中的姓名完全匹配
3. 日期格式必须与原始表格中的格式一致
4. 建议在处理前备份原始文件 