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

        public MainForm()
        {
            Text = "LedgerFrame";
            WindowState = FormWindowState.Maximized;
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
        }

        void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                status.Text = "正在加载模型…";
                var sw = System.Diagnostics.Stopwatch.StartNew();
                // 开发期:从 exe 同目录读模型;打包后改内嵌加密资源
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "model_syncf.xlsx");
                if (!File.Exists(path))
                    path = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\model_syncf.xlsx"));
                spreadsheet.Open(path);
                sw.Stop();
                // 全表启用计算引擎(懒计算,显示到哪算哪)
                foreach (IWorksheet sh in spreadsheet.Workbook.Worksheets)
                    sh.EnableSheetCalculations();
                status.Text = string.Format("就绪（模型加载 {0:F1} 秒）", sw.Elapsed.TotalSeconds);
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
                    var r = Injector.Inject(spreadsheet.Workbook, dlg.FileName);
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
