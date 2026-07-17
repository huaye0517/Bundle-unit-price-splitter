# 组合装单价拆分工具

Windows 原生桌面工具：上传组合装拆分占比表和销售单原表，一键生成带 AA、AB、AC、AD、AH 拆分数据的结果工作簿。计算引擎已经内嵌在单个 EXE 中。

## 计算规则

- AA：使用销售表货品编号，在旧版占比表 `Sheet2` 中匹配“单品编号/单价”，或在新版导出表（如 `sheet1`、`总表`）中匹配“编号/执行价格”；合并文件中重复出现的表头会自动跳过，未匹配填 0。
- 原“金额”为 0 的行不参与占比和金额分摊，AA、AD、AH 均填 0。
- 以销售表 `Sheet1` 的 L 列“网店订单号”匹配 N 列“订单金额”，并按网店订单分组拆分。
- AB：`AA / 网店订单内 AA 合计 / 数量`。
- AC：`Sheet1 订单金额 × AB`。
- AD：`AC × 数量`。
- AH：与 AD 逐行一致。
- 整个网店订单都未匹配占比表时，按各行原金额占比分摊 Sheet1 订单金额；原金额为 0 的行仍不参与。
- 所有列写入计算值，不写外部引用公式，不额外舍入。
- Sheet1“订单金额”的外部公式会使用源文件中的已计算结果固化为数值，避免保存后显示为空。

## 运行源码

```powershell
python splitter_worker.py 占比表.xlsx 销售单.xlsx 输出.xlsx
```

## 测试与打包

```powershell
.\build.ps1
```

打包完成后，程序位于 `outputs\组合装单价拆分工具.exe`，可直接复制到其他 Windows 电脑运行。

如需指定打包使用的 Python，可先设置 `BUNDLE_SPLITTER_PYTHON` 环境变量。真实样例回归测试可通过 `BUNDLE_SPLITTER_RATIO_SAMPLE` 和 `BUNDLE_SPLITTER_SALES_SAMPLE` 指定本机文件，不会把业务表上传到仓库。
