using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Syncfusion.Windows.Forms.Spreadsheet;
using Syncfusion.XlsIO;

namespace LedgerFrame
{
    public class MainForm : Form
    {
        readonly Spreadsheet spreadsheet;
        readonly ToolStripStatusLabel status;
        System.Collections.Generic.List<ValidationFixer.Entry> dvEntries =
            new System.Collections.Generic.List<ValidationFixer.Entry>();

        public MainForm()
        {
            Text = "LedgerFrame";
            // net462 + PerMonitorV2 下 WindowState.Maximized 会把客户区算高,底部被裁掉一截;
            // 直接用工作区大小代替最大化
            WindowState = FormWindowState.Normal;
            StartPosition = FormStartPosition.Manual;
            Bounds = Screen.PrimaryScreen.WorkingArea;
            Font = new Font("Microsoft YaHei UI", 9F);

            var tools = new ToolStrip();
            var btnUpdate = new ToolStripButton("更新数据");
            var btnExport = new ToolStripButton("导出打印");
            tools.Items.Add(btnUpdate);
            tools.Items.Add(btnExport);
            tools.Dock = DockStyle.Top;

            var strip = new StatusStrip();
            status = new ToolStripStatusLabel("就绪");
            strip.Items.Add(status);
            strip.Dock = DockStyle.Bottom;

            spreadsheet = new Spreadsheet { Dock = DockStyle.Fill };

            Controls.Add(spreadsheet);
            Controls.Add(tools);
            Controls.Add(strip);

            Load += MainForm_Load;
            btnUpdate.Click += BtnUpdate_Click;
            btnExport.Click += (s, e) =>
                MessageBox.Show(this, "打印/导出功能后期再加（外挂模块）。", "提示");

            // 状态反馈:切表/编辑后更新状态栏
            spreadsheet.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == "ActiveSheet")
                {
                    HookGrid();
                    status.Text = "当前表：" + spreadsheet.ActiveSheet.Name;
                }
            };
            HookGrid();
        }

        readonly System.Collections.Generic.HashSet<object> hookedGrids =
            new System.Collections.Generic.HashSet<object>();

        void HookGrid()
        {
            var g = spreadsheet.ActiveGrid;
            if (g == null || hookedGrids.Contains(g)) return;
            hookedGrids.Add(g);
            g.CurrentCellEndEdit += (s, e) =>
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                try
                {
                    ValidationFixer.OnCellEdited(spreadsheet.Workbook, dvEntries,
                        spreadsheet.ActiveSheet, g.CurrentCell.RowIndex, g.CurrentCell.ColumnIndex);
                    g.InvalidateCells();
                }
                catch { }
                sw.Stop();
                status.Text = string.Format("已编辑 R{0}C{1}，刷新 {2} ms",
                    g.CurrentCell.RowIndex, g.CurrentCell.ColumnIndex, sw.ElapsedMilliseconds);
            };
        }

        // 读取内嵌加密资源并解密;找不到资源返回 null(开发期回退到明文文件)
        static byte[] LoadEmbedded(string resName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var s = asm.GetManifestResourceStream(resName))
            {
                if (s == null) return null;
                var buf = new byte[s.Length];
                s.Read(buf, 0, buf.Length);
                return KeyHolder.Decrypt(buf);
            }
        }

        void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                status.Text = "正在加载模型…";
                var sw = System.Diagnostics.Stopwatch.StartNew();
                byte[] model = LoadEmbedded("LedgerFrame.model.dat");
                if (model != null)
                {
                    spreadsheet.Open(new MemoryStream(model));
                }
                else
                {
                    // 开发期:明文模型;仓库里只有 samples/TreeVolumeSample.xlsx
                    string path = FindDevFile("model_syncf.xlsx") ?? FindDevFile("TreeVolumeSample.xlsx");
                    if (path == null)
                    {
                        status.Text = "未找到模型文件";
                        return;
                    }
                    spreadsheet.Open(path);
                }
                sw.Stop();
                // 全表启用计算引擎(懒计算,显示到哪算哪)
                foreach (IWorksheet sh in spreadsheet.Workbook.Worksheets)
                    sh.EnableSheetCalculations();
                // INDIRECT 动态验证解析成静态区域,否则控件不显示下拉
                var dvSw = System.Diagnostics.Stopwatch.StartNew();
                dvEntries = ValidationFixer.ResolveAll(spreadsheet.Workbook);
                dvSw.Stop();
                status.Text = string.Format("就绪：{0} 张表（加载 {1:F1} 秒，验证解析 {2} 处/{3:F1} 秒）",
                    spreadsheet.Workbook.Worksheets.Count, sw.Elapsed.TotalSeconds,
                    dvEntries.Count, dvSw.Elapsed.TotalSeconds);
                GotoDebugCell();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "模型加载失败：\r\n" + ex, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                status.Text = "加载失败";
            }
        }

        void BtnUpdate_Click(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "选择采伐表数据文件";
                dlg.Filter = "Excel 文件|*.xlsx;*.xlsm;*.xls";
                if (dlg.ShowDialog(this) != DialogResult.OK)
                    return;
                try
                {
                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    byte[] mapBytes = LoadEmbedded("LedgerFrame.mapping.dat");
                    string mapJson;
                    if (mapBytes != null)
                        mapJson = System.Text.Encoding.UTF8.GetString(mapBytes);
                    else
                    {
                        string mapPath = FindDevFile("extlink_mapping.json");
                        if (mapPath == null)
                        {
                            MessageBox.Show(this, "此构建未包含数据注入映射。", "提示");
                            return;
                        }
                        mapJson = File.ReadAllText(mapPath);
                    }
                    var r = Injector.Inject(spreadsheet.Workbook, dlg.FileName, mapJson);
                    RefreshAllGrids();
                    sw.Stop();
                    status.Text = string.Format(
                        "更新数据完成：注入 {0} 格，跳过 {1} 格（源文件缺对应表），耗时 {2:F1} 秒",
                        r.Injected, r.Skipped, sw.Elapsed.TotalSeconds);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "更新数据失败：\r\n" + ex.Message, "错误",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        static string FindDevFile(string name)
        {
            string p = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, name);
            if (File.Exists(p)) return p;
            p = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\" + name));
            if (File.Exists(p)) return p;
            return null;
        }

        // 调试参数 --goto=表名!单元格,用于自动化截图验证
        void GotoDebugCell()
        {
            foreach (var a in Environment.GetCommandLineArgs())
            {
                if (!a.StartsWith("--goto=")) continue;
                var arg = a.Substring(7);
                int bang = arg.LastIndexOf('!');
                if (bang <= 0) continue;
                string sheetName = arg.Substring(0, bang);
                string addr = arg.Substring(bang + 1);
                try
                {
                    spreadsheet.SetActiveSheet(sheetName);
                    var cell = spreadsheet.Workbook.Worksheets[sheetName].Range[addr];
                    var g = spreadsheet.ActiveGrid;
                    g.CurrentCell.MoveCurrentCell(cell.Row, cell.Column);
                    g.ScrollInView(new Syncfusion.Windows.Forms.CellGrid.ScrollAxis.RowColumnIndex(cell.Row, cell.Column));
                    g.Focus();
                }
                catch (Exception ex) { status.Text = "goto 失败:" + ex.Message; }
            }
        }

        void RefreshAllGrids()
        {
            try
            {
                foreach (var g in spreadsheet.GridCollection.Values)
                    g.InvalidateCells();
            }
            catch
            {
                spreadsheet.Invalidate(true);
            }
        }
    }
}
