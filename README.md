# 组合装单价拆分工具

Windows 原生桌面工具：上传组合装拆分占比表和销售单原表，一键生成带 AA、AB、AC、AD、AH 拆分数据的结果工作簿。计算引擎已经内嵌在单个 EXE 中。

## 计算规则

- AA：使用销售表货品编号，在占比表 `Sheet2` 的 B:E 中做首条精确匹配，返回 E 列单价；未匹配填 0。
- 按“网店订单号”独立分组，每个网店订单以 `Sheet1` 的 N 列“应收合计”为分摊目标（同一订单只取该列的非空值）。
- AB：`AA / 网店订单内 AA 合计 / 数量`。
- AC：`网店订单应收合计 × AB`。
- AD：`AC × 数量`。
- AH：等于 AD。
- 整个网店订单都未匹配占比表时，按各行原金额占比分摊 N 列“应收合计”，确保订单级 AH 合计仍与应收一致。
- 所有列写入计算值，不写外部引用公式，不额外舍入。

## 运行源码

```powershell
python splitter_worker.py 占比表.xlsx 销售单.xlsx 输出.xlsx
```

## 测试与打包

```powershell
.\build.ps1
```

打包完成后，程序位于 `outputs\BundleUnitPriceSplitter_v3.exe`，可直接复制到其他 Windows 电脑运行。

如需指定打包使用的 Python，可先设置 `BUNDLE_SPLITTER_PYTHON` 环境变量。真实样例回归测试可通过 `BUNDLE_SPLITTER_RATIO_SAMPLE` 和 `BUNDLE_SPLITTER_SALES_SAMPLE` 指定本机文件，不会把业务表上传到仓库。
