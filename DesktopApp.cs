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
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("SplitterWorker.exe"))
            {
                if (stream == null) throw new InvalidOperationException("Embedded worker is missing.");
                File.WriteAllText(args[1], "ok embedded_worker_bytes=" + stream.Length, Encoding.UTF8);
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
    private readonly ProgressBar progressBar;
    private readonly Button generateButton;
    private string ratioPath = "";
    private string salesPath = "";
    private string outputPath = "";
    private string workerPath = "";

    internal SplitterForm()
    {
        Text = "组合装单价拆分工具";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(780, 650);
        ClientSize = new Size(900, 690);
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
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 118));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 12));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 118));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 146));
        content.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
        content.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        Controls.Add(content);

        content.Controls.Add(MakeLabel("BUNDLE / UNIT LEDGER", new Font("Consolas", 9F, FontStyle.Bold), Teal), 0, 0);
        content.Controls.Add(MakeLabel("组合装单价拆分", new Font("Microsoft YaHei UI", 24F, FontStyle.Bold), Ink), 0, 1);
        content.Controls.Add(MakeLabel("读取 Sheet1 订单金额，按网店订单完成 AA、AB、AC、AD、AH 计算。", BodyFont, Muted), 0, 2);

        ratioPathLabel = MakePathLabel("尚未选择文件");
        content.Controls.Add(MakeFileCard("组合装拆分占比表", "自动识别旧版 Sheet2，或新版 sheet1/总表：编号、执行价格", ratioPathLabel, ChooseRatio), 0, 3);
        salesPathLabel = MakePathLabel("尚未选择文件");
        content.Controls.Add(MakeFileCard("销售单原表", "按网店订单号分组，引用 Sheet1 的订单金额", salesPathLabel, ChooseSales), 0, 5);

        var ledger = new Panel { Dock = DockStyle.Fill, BackColor = Paper, Padding = new Padding(0, 14, 0, 8) };
        var ledgerTitle = MakeLabel("金额核对", new Font("Microsoft YaHei UI", 11F, FontStyle.Bold), Ink);
        ledgerTitle.Location = new Point(0, 12); ledgerTitle.AutoSize = true;
        balanceLabel = MakeLabel("等待计算 · 订单金额 —  拆分 —  差额 —", new Font("Consolas", 10F, FontStyle.Bold), Teal);
        balanceLabel.BackColor = TealSoft; balanceLabel.Location = new Point(0, 43); balanceLabel.Size = new Size(816, 38);
        balanceLabel.Padding = new Padding(12, 9, 12, 8); balanceLabel.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        progressBar = new ProgressBar { Location = new Point(0, 88), Size = new Size(816, 10), Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top, Style = ProgressBarStyle.Continuous };
        statusLabel = MakeLabel("选择两份 Excel 后即可开始", BodyFont, Muted);
        statusLabel.Location = new Point(0, 108); statusLabel.AutoSize = true;
        ledger.Controls.Add(ledgerTitle); ledger.Controls.Add(balanceLabel); ledger.Controls.Add(progressBar); ledger.Controls.Add(statusLabel);
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

        try { workerPath = ExtractWorker(); }
        catch (Exception ex) { MessageBox.Show("无法准备计算引擎：" + ex.Message, "启动失败", MessageBoxButtons.OK, MessageBoxIcon.Error); generateButton.Enabled = false; }
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

    private void ShowSelectedFile(Label label, string path)
    {
        label.Text = "✓  已选择  " + Path.GetFileName(path);
        label.ForeColor = Teal;
        label.BackColor = TealSoft;
        label.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold);
        fileToolTip.SetToolTip(label, path);
    }

    private Panel MakeFileCard(string title, string hint, Label pathLabel, EventHandler click)
    {
        var panel = new Panel { Dock = DockStyle.Fill, BackColor = Surface, Padding = new Padding(20, 14, 20, 14) };
        var grid = new TableLayoutPanel { Dock = DockStyle.Fill, BackColor = Surface, ColumnCount = 2, RowCount = 3 };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 125));
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 27)); grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 27)); grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        var titleLabel = MakeLabel(title, new Font("Microsoft YaHei UI", 12F, FontStyle.Bold), Ink); titleLabel.BackColor = Surface;
        var hintLabel = MakeLabel(hint, BodyFont, Muted); hintLabel.BackColor = Surface;
        var button = MakeButton("选择文件", false, click); button.Dock = DockStyle.Fill; button.Margin = new Padding(10, 4, 0, 4);
        grid.Controls.Add(titleLabel, 0, 0); grid.Controls.Add(hintLabel, 0, 1); grid.Controls.Add(pathLabel, 0, 2); grid.Controls.Add(button, 1, 0); grid.SetRowSpan(button, 3);
        panel.Controls.Add(grid); return panel;
    }

    private Button MakeButton(string text, bool primary, EventHandler click)
    {
        var button = new Button { Text = text, FlatStyle = FlatStyle.Flat, Font = new Font("Microsoft YaHei UI", 10F, primary ? FontStyle.Bold : FontStyle.Regular), Cursor = Cursors.Hand, Size = new Size(140, 42) };
        button.FlatAppearance.BorderSize = 0; button.BackColor = primary ? Blue : Color.FromArgb(234, 240, 242); button.ForeColor = primary ? Color.White : Blue; button.Click += click;
        return button;
    }

    private void ChooseRatio(object sender, EventArgs e)
    {
        using (var dialog = new OpenFileDialog { Title = "选择组合装拆分占比表", Filter = "Excel 工作簿 (*.xlsx)|*.xlsx" })
            if (dialog.ShowDialog(this) == DialogResult.OK) { ratioPath = dialog.FileName; ShowSelectedFile(ratioPathLabel, ratioPath); UpdateReady(); }
    }

    private void ChooseSales(object sender, EventArgs e)
    {
        using (var dialog = new OpenFileDialog { Title = "选择销售单原表", Filter = "Excel 工作簿 (*.xlsx)|*.xlsx" })
            if (dialog.ShowDialog(this) == DialogResult.OK) {
                salesPath = dialog.FileName; ShowSelectedFile(salesPathLabel, salesPath);
                outputPath = Path.Combine(Path.GetDirectoryName(salesPath), Path.GetFileNameWithoutExtension(salesPath) + "_已拆单价_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
                outputPathLabel.Text = "输出：" + outputPath; UpdateReady();
            }
    }

    private void ChooseOutput(object sender, EventArgs e)
    {
        using (var dialog = new SaveFileDialog { Title = "设置输出文件", Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", DefaultExt = "xlsx", FileName = string.IsNullOrEmpty(outputPath) ? "已拆单价表.xlsx" : Path.GetFileName(outputPath), InitialDirectory = string.IsNullOrEmpty(outputPath) ? "" : Path.GetDirectoryName(outputPath) })
            if (dialog.ShowDialog(this) == DialogResult.OK) { outputPath = dialog.FileName; outputPathLabel.Text = "输出：" + outputPath; }
    }

    private void UpdateReady()
    {
        if (!string.IsNullOrEmpty(ratioPath) && !string.IsNullOrEmpty(salesPath)) {
            statusLabel.Text = "两份文件已就绪";
            balanceLabel.Text = "准备完成 · 点击“一键拆分并生成 Excel”";
        }
    }

    private async void Generate(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(ratioPath) || string.IsNullOrEmpty(salesPath)) { MessageBox.Show("请先选择组合装拆分占比表和销售单原表。", "还缺文件", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        generateButton.Enabled = false; progressBar.Style = ProgressBarStyle.Continuous; progressBar.Value = 0;
        statusLabel.Text = "正在读取数据并按网店订单计算"; balanceLabel.Text = "正在核对 · 订单金额计算中  拆分计算中  差额计算中";
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
            var message = string.Format("结果已保存：\n{0}\n\n销售单组：{1}\n匹配行：{2}\n未匹配普通商品：{3}\n异常订单：{4}\n\n是否打开结果所在文件夹？", outputPath, orders, matched, unmatched, exceptional);
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
        var info = new ProcessStartInfo { FileName = workerPath, Arguments = Quote(ratioPath) + " " + Quote(salesPath) + " " + Quote(outputPath), UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true, StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8 };
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
                    if (separator > 0 && int.TryParse(value.Substring(0, separator), out percent)) onProgress(percent, Decode(value.Substring(separator + 1)));
                } else values[key] = value;
            }
            process.WaitForExit();
            string stderr = errorTask.Result;
            if (process.ExitCode != 0 && !values.ContainsKey("message_b64")) throw new Exception(string.IsNullOrWhiteSpace(stderr) ? "计算引擎异常退出。" : stderr.Trim());
            return values;
        }
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
