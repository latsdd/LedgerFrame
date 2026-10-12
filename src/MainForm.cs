using System;
using System.Collections.Generic;
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
        string vmJson;   // validation_map.json 内容(内嵌或开发期文件)

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
                    EnsureValidationsResolved();
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
            // Excel 习惯: 点击列表验证格 -> 直接进编辑态显示下拉箭头
            // (该控件的组合框渲染器只在编辑态出现)
            g.CellClick += (s, e) =>
            {
                try
                {
                    var cell = spreadsheet.ActiveSheet.Range[e.RowIndex, e.ColumnIndex];
                    if (!cell.HasDataValidation) return;
                    var dv = cell.DataValidation;
                    if (dv == null || dv.AllowType != ExcelDataType.User) return;
                    var cc = g.CurrentCell;
                    if (cc.IsEditing) return;
                    cc.BeginEdit(true);
                }
                catch { }
            };
        }

        // 控件对未激活表不建 grid, 验证访问会 NRE; 推迟到表第一次激活时重建联动下拉
        readonly HashSet<string> dvDoneSheets = new HashSet<string>();

        void EnsureValidationsResolved()
        {
            if (vmJson == null || spreadsheet.ActiveSheet == null) return;
            var sh = spreadsheet.ActiveSheet;
            if (dvDoneSheets.Contains(sh.Name)) return;
            dvDoneSheets.Add(sh.Name);
            try
            {
                var list = ValidationFixer.LoadMapForSheet(spreadsheet.Workbook, vmJson, sh.Name);
                dvEntries.AddRange(list);
                if (list.Count > 0)
                    DvLog("resolved " + sh.Name + ": " + list.Count + " 组, 首格 f1="
                        + list[0].Ranges[0].DataValidation.FirstFormula);
            }
            catch (Exception ex) { DvLog("resolve " + sh.Name + " EX: " + ex.Message); }
        }

        static void DvLog(string line)
        {
            try { System.IO.File.AppendAllText(@"C:\tmp\lf_dv.log", line + "\n"); } catch { }
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
                DvLog("--- load ---");
                // 联动下拉映射只加载文本, 重建推迟到各表首次激活(控件对未激活表不建 grid)
                try
                {
                    byte[] vmBytes = LoadEmbedded("LedgerFrame.valmap.dat");
                    if (vmBytes != null)
                        vmJson = System.Text.Encoding.UTF8.GetString(vmBytes);
                    else
                    {
                        string mapPath = FindDevFile("validation_map.json");
                        if (mapPath != null) vmJson = System.IO.File.ReadAllText(mapPath);
                    }
                }
                catch (Exception exR) { DvLog("vmJson EX: " + exR.Message); }
                // 全表启用计算引擎(懒计算,显示到哪算哪)
                foreach (IWorksheet sh in spreadsheet.Workbook.Worksheets)
                    sh.EnableSheetCalculations();
                status.Text = string.Format("就绪：{0} 张表（加载 {1:F1} 秒）",
                    spreadsheet.Workbook.Worksheets.Count, sw.Elapsed.TotalSeconds);
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

        // 调试参数 --goto=表名或序号!单元格,用于自动化截图验证
        // 在 Shown 后延迟执行:Load 阶段控件尚未完成布局,切表会被重置
        void GotoDebugCell()
        {
            string target = null;
            foreach (var a in Environment.GetCommandLineArgs())
                if (a.StartsWith("--goto=")) target = a.Substring(7);
            if (target == null) return;
            var timer = new Timer { Interval = 1000 };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                try
                {
                    int bang = target.LastIndexOf('!');
                    string sheetName = target.Substring(0, bang);
                    string addr = target.Substring(bang + 1);
                    int idx;
                    if (int.TryParse(sheetName, out idx))
                        spreadsheet.SetActiveSheet(spreadsheet.Workbook.Worksheets[idx].Name);
                    else
                        spreadsheet.SetActiveSheet(sheetName);
                    var cell = spreadsheet.ActiveSheet.Range[addr];
                    var g = spreadsheet.ActiveGrid;
                    g.CurrentCell.MoveCurrentCell(cell.Row, cell.Column);
                    g.ScrollInView(new Syncfusion.Windows.Forms.CellGrid.ScrollAxis.RowColumnIndex(cell.Row, cell.Column));
                    g.Focus();
                    // 诊断: 网格渲染器列表 + 该格验证状态
                    string diag = "";
                    try
                    {
                        var keys = new System.Text.StringBuilder();
                        foreach (var k in g.CellRenderers.Keys) keys.Append(k).Append('|');
                        diag = " rend=" + keys
                            + " B3HasDV=" + spreadsheet.ActiveSheet.Range["B3"].HasDataValidation;
                    }
                    catch (Exception dx) { diag = " diagEX=" + dx.Message; }
                    System.IO.File.WriteAllText(@"C:\tmp\lf_goto.log",
                        "OK " + spreadsheet.ActiveSheet.Name + " " + addr + diag);
                    // 诊断2: 进入编辑态, 看组合框编辑器是否出现
                    bool wantEdit = false;
                    foreach (var aa in Environment.GetCommandLineArgs()) if (aa == "--edit") wantEdit = true;
                    if (wantEdit)
                    {
                        try
                        {
                            g.CurrentCell.BeginEdit(true);
                            System.IO.File.AppendAllText(@"C:\tmp\lf_goto.log",
                                "\nBeginEdit: IsEditing=" + g.CurrentCell.IsEditing);
                            SendKeys.SendWait("{F4}");
                            System.IO.File.AppendAllText(@"C:\tmp\lf_goto.log", "\nF4 sent");
                        }
                        catch (Exception be) { System.IO.File.AppendAllText(@"C:\tmp\lf_goto.log", "\nBeginEdit EX: " + be.Message); }
                    }
                }
                catch (Exception ex)
                {
                    System.IO.File.WriteAllText(@"C:\tmp\lf_goto.log", "FAIL " + ex);
                }
            };
            timer.Start();
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
