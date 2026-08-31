using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        if (args.Length == 2 && args[0] == "--self-test")
        {
            using (var worker = Assembly.GetExecutingAssembly().GetManifestResourceStream("SplitterWorker.exe"))
            using (var ratio = Assembly.GetExecutingAssembly().GetManifestResourceStream("DefaultRatioData.xlsx"))
            {
                if (worker == null) throw new InvalidOperationException("Embedded worker is missing.");
                if (ratio == null) throw new InvalidOperationException("Embedded ratio data is missing.");
                File.WriteAllText(args[1], "ok embedded_worker_bytes=" + worker.Length + " embedded_ratio_bytes=" + ratio.Length, Encoding.UTF8);
            }
            return;
        }
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new SplitterForm());
    }
}

internal sealed class SplitterForm : Form
{
    private readonly Color Paper = Color.FromArgb(243, 246, 247);
    private readonly Color Surface = Color.White;
    private readonly Color Ink = Color.FromArgb(24, 49, 61);
    private readonly Color Muted = Color.FromArgb(101, 122, 132);
    private readonly Color Blue = Color.FromArgb(36, 92, 115);
    private readonly Color Teal = Color.FromArgb(20, 129, 109);
    private readonly Color TealSoft = Color.FromArgb(221, 241, 236);
    private readonly Font BodyFont = new Font("Microsoft YaHei UI", 10F);
    private readonly Font MonoFont = new Font("Consolas", 9F);
    private readonly ToolTip fileToolTip = new ToolTip();

    private readonly Label ratioPathLabel;
    private readonly Label salesPathLabel;
    private readonly Label outputPathLabel;
    private readonly Label statusLabel;
    private readonly Label balanceLabel;
    private readonly RadioButton appendRatioRadio;
    private readonly RadioButton replaceRatioRadio;
    private readonly RadioButton feeModeRadio;
    private readonly RadioButton noFeeModeRadio;
    private readonly Button ratioUpdateButton;
    private readonly ProgressBar progressBar;
    private readonly Button generateButton;
    private string ratioPath = "";
    private string salesPath = "";
    private string outputPath = "";
    private string workerPath = "";

    internal SplitterForm()
    {
        Text = "组合装单价拆分工具";
        Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(780, 690);
        ClientSize = new Size(900, 720);
        BackColor = Paper;
        Font = BodyFont;
        AutoScaleMode = AutoScaleMode.Dpi;

        var content = new TableLayoutPanel {
            Dock = DockStyle.Fill, Padding = new Padding(42, 30, 42, 26),
            BackColor = Paper, ColumnCount = 1, RowCount = 9
        };
        content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 136));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 12));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 142));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 146));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
        content.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        Controls.Add(content);

        content.Controls.Add(MakeLabel("BUNDLE / UNIT LEDGER", new Font("Consolas", 9F, FontStyle.Bold), Teal), 0, 0);
        content.Controls.Add(MakeLabel("组合装单价拆分", new Font("Microsoft YaHei UI", 24F, FontStyle.Bold), Ink), 0, 1);
        content.Controls.Add(MakeLabel("按列头自动识别订单、物流、应收、货品、数量、单价、金额和备注。", BodyFont, Muted), 0, 2);

        ratioPathLabel = MakePathLabel("正在准备内置基础数据…");
        appendRatioRadio = MakeRadio("新增", true);
        replaceRatioRadio = MakeRadio("覆盖", false);
        ratioUpdateButton = MakeButton("选择更新文件", false, ChooseRatioUpdate);
        content.Controls.Add(MakeRatioDataCard(), 0, 3);
        feeModeRadio = MakeRadio("有手续费", true);
        noFeeModeRadio = MakeRadio("无手续费", false);
        feeModeRadio.CheckedChanged += FeeModeChanged;
        noFeeModeRadio.CheckedChanged += FeeModeChanged;
        salesPathLabel = MakePathLabel("尚未选择文件 · 也可将 .xlsx 拖到这里");
        var salesCard = MakeSalesFileCard();
        EnableExcelDrop(salesCard, DropSales);
        content.Controls.Add(salesCard, 0, 5);

        var ledger = new TableLayoutPanel {
            Dock = DockStyle.Fill, BackColor = Paper, Padding = new Padding(0, 10, 0, 6),
            ColumnCount = 1, RowCount = 4, Margin = new Padding(0)
        };
        ledger.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        ledger.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        ledger.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
        ledger.RowStyles.Add(new RowStyle(SizeType.Absolute, 18));
        ledger.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
        var ledgerTitle = MakeLabel("金额核对", new Font("Microsoft YaHei UI", 11F, FontStyle.Bold), Ink);
        ledgerTitle.Margin = new Padding(0);
        balanceLabel = MakeLabel("等待计算 · 应收合计 —  拆分 —  差额 —", new Font("Consolas", 10F, FontStyle.Bold), Teal);
        balanceLabel.BackColor = TealSoft; balanceLabel.Padding = new Padding(12, 0, 12, 0); balanceLabel.Margin = new Padding(0);
        progressBar = new ProgressBar { Dock = DockStyle.Fill, Margin = new Padding(0, 4, 0, 4), Style = ProgressBarStyle.Continuous };
        statusLabel = MakeLabel("基础库已内置，拖入销售单后开始拆分", BodyFont, Muted);
        statusLabel.Margin = new Padding(0);
        ledger.Controls.Add(ledgerTitle, 0, 0);
        ledger.Controls.Add(balanceLabel, 0, 1);
        ledger.Controls.Add(progressBar, 0, 2);
        ledger.Controls.Add(statusLabel, 0, 3);
        content.Controls.Add(ledger, 0, 6);

        var actions = new TableLayoutPanel { Dock = DockStyle.Fill, BackColor = Paper, ColumnCount = 2, RowCount = 1 };
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        actions.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        var outputButton = MakeButton("设置输出位置", false, ChooseOutput);
        outputButton.Anchor = AnchorStyles.Left;
        generateButton = MakeButton("一键拆分并生成 Excel", true, Generate);
        generateButton.Dock = DockStyle.Fill;
        generateButton.Margin = new Padding(54, 6, 0, 8);
        actions.Controls.Add(outputButton, 0, 0); actions.Controls.Add(generateButton, 1, 0);
        content.Controls.Add(actions, 0, 7);

        outputPathLabel = MakeLabel("输出位置将在选择销售单后自动生成", BodyFont, Muted);
        outputPathLabel.Dock = DockStyle.Top; outputPathLabel.AutoEllipsis = true;
        content.Controls.Add(outputPathLabel, 0, 8);

        try {
            workerPath = ExtractWorker();
            ratioPath = EnsureRatioData();
            RefreshRatioInfo();
        }
        catch (Exception ex) { MessageBox.Show("无法准备内置基础数据：" + ex.Message, "启动失败", MessageBoxButtons.OK, MessageBoxIcon.Error); generateButton.Enabled = false; ratioUpdateButton.Enabled = false; }
    }

    private Label MakeLabel(string text, Font font, Color color)
    {
        return new Label { Text = text, Font = font, ForeColor = color, BackColor = Color.Transparent, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
    }

    private Label MakePathLabel(string text)
    {
        return new Label {
            Text = text, Font = new Font("Microsoft YaHei UI", 9.5F), ForeColor = Muted,
            BackColor = Color.FromArgb(246, 249, 250), AutoEllipsis = true,
            Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(10, 0, 10, 0), BorderStyle = BorderStyle.FixedSingle
        };
    }

    private RadioButton MakeRadio(string text, bool selected)
    {
        return new RadioButton {
            Text = text, Checked = selected, AutoSize = true, Font = new Font("Microsoft YaHei UI", 9F),
            ForeColor = Ink, BackColor = Surface, Margin = new Padding(0, 4, 14, 0)
        };
    }

    private void ShowSelectedFile(Label label, string path)
    {
        label.Text = "✓  已选择  " + Path.GetFileName(path);
        label.ForeColor = Teal;
        label.BackColor = TealSoft;
        label.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold);
        fileToolTip.SetToolTip(label, path);
    }

    private Panel MakeSalesFileCard()
    {
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = Surface, Padding = new Padding(20, 12, 20, 12) };
        var grid = new TableLayoutPanel { Dock = DockStyle.Fill, BackColor = Surface, ColumnCount = 2, RowCount = 4 };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 125));
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 27)); grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 29)); grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        var titleLabel = MakeLabel("销售单原表", new Font("Microsoft YaHei UI", 12F, FontStyle.Bold), Ink); titleLabel.BackColor = Surface;
        var hintLabel = MakeLabel("选择拆分模式后上传销售单；列头会自动识别", BodyFont, Muted); hintLabel.BackColor = Surface;
        var modes = new FlowLayoutPanel { Dock = DockStyle.Fill, BackColor = Surface, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0) };
        var modeLabel = MakeLabel("拆分模式：", new Font("Microsoft YaHei UI", 9F, FontStyle.Bold), Muted);
        modeLabel.AutoSize = true; modeLabel.Dock = DockStyle.None; modeLabel.Margin = new Padding(0, 5, 8, 0);
        modes.Controls.Add(modeLabel); modes.Controls.Add(feeModeRadio); modes.Controls.Add(noFeeModeRadio);
        var button = MakeButton("选择文件", false, ChooseSales); button.Dock = DockStyle.Fill; button.Margin = new Padding(10, 4, 0, 4);
        grid.Controls.Add(titleLabel, 0, 0); grid.Controls.Add(hintLabel, 0, 1); grid.Controls.Add(modes, 0, 2); grid.Controls.Add(salesPathLabel, 0, 3);
        grid.Controls.Add(button, 1, 0); grid.SetRowSpan(button, 4);
        panel.Controls.Add(grid);
        return panel;
    }

    private Panel MakeRatioDataCard()
    {
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = Surface, Padding = new Padding(20, 14, 20, 14) };
        var grid = new TableLayoutPanel { Dock = DockStyle.Fill, BackColor = Surface, ColumnCount = 2, RowCount = 3 };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 27)); grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 27)); grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        var titleLabel = MakeLabel("组合装基础数据", new Font("Microsoft YaHei UI", 12F, FontStyle.Bold), Ink); titleLabel.BackColor = Surface;
        var hintLabel = MakeLabel("已内置《组合装及子件合并数据》；拖入 .xlsx 可更新", BodyFont, Muted); hintLabel.BackColor = Surface;
        var controls = new TableLayoutPanel { Dock = DockStyle.Fill, BackColor = Surface, ColumnCount = 1, RowCount = 2, Margin = new Padding(10, 0, 0, 0) };
        controls.RowStyles.Add(new RowStyle(SizeType.Absolute, 32)); controls.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        var modes = new FlowLayoutPanel { Dock = DockStyle.Fill, BackColor = Surface, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0) };
        modes.Controls.Add(appendRatioRadio); modes.Controls.Add(replaceRatioRadio);
        ratioUpdateButton.Dock = DockStyle.Fill; ratioUpdateButton.Margin = new Padding(0, 3, 0, 0);
        controls.Controls.Add(modes, 0, 0); controls.Controls.Add(ratioUpdateButton, 0, 1);
        grid.Controls.Add(titleLabel, 0, 0); grid.Controls.Add(hintLabel, 0, 1); grid.Controls.Add(ratioPathLabel, 0, 2);
        grid.Controls.Add(controls, 1, 0); grid.SetRowSpan(controls, 3);
        panel.Controls.Add(grid);
        EnableExcelDrop(panel, DropRatio);
        return panel;
    }

    private Button MakeButton(string text, bool primary, EventHandler click)
    {
        var button = new Button { Text = text, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft YaHei UI", 10F, primary ? FontStyle.Bold : FontStyle.Regular), Cursor = Cursors.Hand, Size = new Size(140, 42) };
        button.FlatAppearance.BorderSize = 0; button.BackColor = primary ? Blue : Color.FromArgb(234, 240, 242); button.ForeColor = primary ? Color.White : Blue; button.Click += click;
        return button;
    }

    private async void ChooseRatioUpdate(object sender, EventArgs e)
    {
        using (var dialog = new OpenFileDialog { Title = "选择基础数据更新文件", Filter = "Excel 工作簿 (*.xlsx)|*.xlsx" })
            if (dialog.ShowDialog(this) == DialogResult.OK) await ImportRatioFile(dialog.FileName);
    }

    private void ChooseSales(object sender, EventArgs e)
    {
        using (var dialog = new OpenFileDialog { Title = "选择销售单原表", Filter = "Excel 工作簿 (*.xlsx)|*.xlsx" })
            if (dialog.ShowDialog(this) == DialogResult.OK) SetSalesPath(dialog.FileName);
    }

    private void EnableExcelDrop(Control control, DragEventHandler dropHandler)
    {
        control.AllowDrop = true;
        control.DragEnter += ExcelDragEnter;
        control.DragDrop += dropHandler;
        foreach (Control child in control.Controls) EnableExcelDrop(child, dropHandler);
    }

    private void ExcelDragEnter(object sender, DragEventArgs e)
    {
        e.Effect = DroppedExcelPath(e) == null ? DragDropEffects.None : DragDropEffects.Copy;
    }

    private static string DroppedExcelPath(DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return null;
        var files = e.Data.GetData(DataFormats.FileDrop) as string[];
        if (files == null || files.Length != 1 || !string.Equals(Path.GetExtension(files[0]), ".xlsx", StringComparison.OrdinalIgnoreCase)) return null;
        return files[0];
    }

    private async void DropRatio(object sender, DragEventArgs e)
    {
        string path = DroppedExcelPath(e);
        if (path != null) await ImportRatioFile(path);
    }

    private void DropSales(object sender, DragEventArgs e)
    {
        string path = DroppedExcelPath(e);
        if (path != null) SetSalesPath(path);
    }

    private void SetSalesPath(string path)
    {
        salesPath = path;
        ShowSelectedFile(salesPathLabel, salesPath);
        outputPath = Path.Combine(Path.GetDirectoryName(salesPath), Path.GetFileNameWithoutExtension(salesPath) + "_已拆单价_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
        outputPathLabel.Text = "输出：" + outputPath;
        UpdateReady();
    }

    private void FeeModeChanged(object sender, EventArgs e)
    {
        if (!feeModeRadio.Checked && !noFeeModeRadio.Checked) return;
        UpdateReady();
    }

    private async Task ImportRatioFile(string path)
    {
        string mode = replaceRatioRadio.Checked ? "replace" : "append";
        if (mode == "replace" && MessageBox.Show(
            "覆盖会用所选文件整体替换当前基础库，是否继续？",
            "确认覆盖基础数据", MessageBoxButtons.YesNo, MessageBoxIcon.Warning
        ) != DialogResult.Yes) return;

        ratioUpdateButton.Enabled = false;
        generateButton.Enabled = false;
        ratioPathLabel.Text = mode == "append" ? "正在新增基础数据…" : "正在覆盖基础数据…";
        try {
            var result = await Task.Run(() => RunWorkerCommand(
                "--update-ratio " + mode + " " + Quote(ratioPath) + " " + Quote(path), null
            ));
            EnsureWorkerSuccess(result);
            int added = ParseInt(result, "added_rows");
            RefreshRatioInfo();
            MessageBox.Show(
                mode == "append" ? string.Format("基础数据已新增 {0} 条组合明细。", added) : "基础数据已完成覆盖。",
                "基础数据已更新", MessageBoxButtons.OK, MessageBoxIcon.Information
            );
        }
        catch (Exception ex) {
            RefreshRatioInfo();
            MessageBox.Show(ex.Message, "基础数据更新失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally {
            ratioUpdateButton.Enabled = true;
            generateButton.Enabled = true;
            UpdateReady();
        }
    }

    private void RefreshRatioInfo()
    {
        var result = RunWorkerCommand("--ratio-info " + Quote(ratioPath), null);
        EnsureWorkerSuccess(result);
        int items = ParseInt(result, "unique_items"), rows = ParseInt(result, "source_rows");
        ratioPathLabel.Text = string.Format("✓  基础库已就绪 · {0} 个子件 · {1} 条组合明细 · 可拖入更新", items, rows);
        ratioPathLabel.ForeColor = Teal;
        ratioPathLabel.BackColor = TealSoft;
        ratioPathLabel.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold);
        fileToolTip.SetToolTip(ratioPathLabel, ratioPath);
    }

    private void ChooseOutput(object sender, EventArgs e)
    {
        using (var dialog = new SaveFileDialog { Title = "设置输出文件", Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", DefaultExt = "xlsx", FileName = string.IsNullOrEmpty(outputPath) ? "已拆单价表.xlsx" : Path.GetFileName(outputPath), InitialDirectory = string.IsNullOrEmpty(outputPath) ? "" : Path.GetDirectoryName(outputPath) })
            if (dialog.ShowDialog(this) == DialogResult.OK) { outputPath = dialog.FileName; outputPathLabel.Text = "输出：" + outputPath; }
    }

    private void UpdateReady()
    {
        if (!string.IsNullOrEmpty(salesPath)) {
            string mode = noFeeModeRadio.Checked ? "无手续费" : "有手续费";
            statusLabel.Text = "销售单已就绪 · " + mode;
            balanceLabel.Text = "准备完成 · " + mode + " · 点击“一键拆分并生成 Excel”";
        } else {
            statusLabel.Text = "基础库已内置，拖入销售单后开始拆分";
        }
    }

    private async void Generate(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(salesPath)) { MessageBox.Show("请先选择或拖入销售单原表。", "还缺销售单", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        generateButton.Enabled = false; progressBar.Style = ProgressBarStyle.Continuous; progressBar.Value = 0;
        string mode = noFeeModeRadio.Checked ? "无手续费" : "有手续费";
        statusLabel.Text = "正在按列头识别数据并计算 · " + mode; balanceLabel.Text = "正在核对 · 应收合计计算中  拆分计算中  差额计算中";
        bool succeeded = false;
        try {
            var result = await Task.Run(() => RunWorker(UpdateProgress));
            if (!result.ContainsKey("status") || result["status"] != "ok") throw new Exception(Decode(result.ContainsKey("message_b64") ? result["message_b64"] : ""));
            int orders = ParseInt(result, "orders"), matched = ParseInt(result, "matched_rows"), unmatched = ParseInt(result, "unmatched_rows"), exceptional = ParseInt(result, "exceptional_orders");
            outputPath = Decode(result["output_path_b64"]);
            statusLabel.Text = string.Format("完成：{0} 个销售单组，{1} 行匹配，{2} 行未匹配", orders, matched, unmatched);
            balanceLabel.Text = exceptional == 0 ? string.Format("核对通过 · 销售单组 {0}  匹配 {1}  普通商品 {2}", orders, matched, unmatched) : string.Format("有 {0} 个异常订单 · 其拆分列已填 0", exceptional);
            progressBar.Value = 100;
            succeeded = true;
            var message = string.Format("结果已保存：\n{0}\n\n拆分模式：{1}\n销售单组：{2}\n匹配行：{3}\n未匹配普通商品：{4}\n异常订单：{5}\n\n是否打开结果所在文件夹？", outputPath, mode, orders, matched, unmatched, exceptional);
            if (MessageBox.Show(message, "拆分完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes) Process.Start("explorer.exe", "/select,\"" + outputPath + "\"");
        } catch (Exception ex) { progressBar.Value = 0; balanceLabel.Text = "核对未通过 · 未生成结果文件"; statusLabel.Text = "生成失败，请按提示检查文件"; MessageBox.Show(ex.Message, "无法生成拆分表", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        finally { if (!succeeded) progressBar.Value = 0; generateButton.Enabled = true; }
    }

    private void UpdateProgress(int percent, string message)
    {
        if (InvokeRequired) { BeginInvoke(new Action<int, string>(UpdateProgress), percent, message); return; }
        int value = Math.Max(0, Math.Min(100, percent));
        progressBar.Value = value;
        statusLabel.Text = message;
        balanceLabel.Text = string.Format("拆分进度 {0}% · {1}", value, message);
    }

    private Dictionary<string, string> RunWorker(Action<int, string> onProgress)
    {
        string mode = noFeeModeRadio.Checked ? "no-fee" : "fee";
        return RunWorkerCommand("--fee-mode " + mode + " " + Quote(ratioPath) + " " + Quote(salesPath) + " " + Quote(outputPath), onProgress);
    }

    private Dictionary<string, string> RunWorkerCommand(string arguments, Action<int, string> onProgress)
    {
        var info = new ProcessStartInfo { FileName = workerPath, Arguments = arguments, UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true, StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8 };
        using (var process = Process.Start(info)) {
            var values = new Dictionary<string, string>();
            var errorTask = Task.Run(() => process.StandardError.ReadToEnd());
            string line;
            while ((line = process.StandardOutput.ReadLine()) != null) {
                int pos = line.IndexOf('=');
                if (pos <= 0) continue;
                string key = line.Substring(0, pos), value = line.Substring(pos + 1);
                if (key == "progress") {
                    int separator = value.IndexOf('|'), percent;
                    if (onProgress != null && separator > 0 && int.TryParse(value.Substring(0, separator), out percent)) onProgress(percent, Decode(value.Substring(separator + 1)));
                } else values[key] = value;
            }
            process.WaitForExit();
            string stderr = errorTask.Result;
            if (process.ExitCode != 0 && !values.ContainsKey("message_b64")) throw new Exception(string.IsNullOrWhiteSpace(stderr) ? "计算引擎异常退出。" : stderr.Trim());
            return values;
        }
    }

    private static void EnsureWorkerSuccess(Dictionary<string, string> result)
    {
        if (!result.ContainsKey("status") || result["status"] != "ok")
            throw new Exception(Decode(result.ContainsKey("message_b64") ? result["message_b64"] : ""));
    }

    private string EnsureRatioData()
    {
        string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "组合装单价拆分工具");
        Directory.CreateDirectory(directory);
        string target = Path.Combine(directory, "组合装及子件合并数据.xlsx");
        if (File.Exists(target)) return target;
        using (var source = Assembly.GetExecutingAssembly().GetManifestResourceStream("DefaultRatioData.xlsx")) {
            if (source == null) throw new InvalidOperationException("内置组合装基础数据缺失。");
            using (var destination = File.Create(target)) source.CopyTo(destination);
        }
        return target;
    }

    private string ExtractWorker()
    {
        string directory = Path.Combine(Path.GetTempPath(), "BundleUnitPriceSplitter"); Directory.CreateDirectory(directory);
        string target = Path.Combine(directory, "SplitterWorker.exe");
        using (var source = Assembly.GetExecutingAssembly().GetManifestResourceStream("SplitterWorker.exe")) {
            if (source == null) throw new InvalidOperationException("内嵌计算引擎缺失。");
            using (var destination = File.Create(target)) source.CopyTo(destination);
        }
        return target;
    }

    private static string Quote(string value) { return "\"" + value.Replace("\"", "\\\"") + "\""; }
    private static string Decode(string value) { return string.IsNullOrEmpty(value) ? "未知错误" : Encoding.UTF8.GetString(Convert.FromBase64String(value)); }
    private static int ParseInt(Dictionary<string, string> values, string key) { int parsed; return values.ContainsKey(key) && int.TryParse(values[key], out parsed) ? parsed : 0; }
}
