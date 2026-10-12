using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Syncfusion.XlsIO;

namespace LedgerFrame
{
    // 外链映射:把采伐表数据文件的值注入模型(原 xlsm 的外部链接公式,补丁1已挖空成注入点)
    static class Injector
    {
        public class Result
        {
            public int Injected;
            public int Skipped;
        }

        class Entry
        {
            public string dst_sheet { get; set; }
            public string dst_cell { get; set; }
            public string src_sheet { get; set; }
            public string src_cell { get; set; }
        }

        // wb: 当前模型(spreadsheet.Workbook);srcPath: 用户选的数据文件;mapJson: 注入点映射
        public static Result Inject(IWorkbook wb, string srcPath, string mapJson)
        {
            var entries = JsonSerializer.Deserialize<List<Entry>>(mapJson);

            var res = new Result();
            using (var engine = new ExcelEngine())
            {
                var srcApp = engine.Excel;
                srcApp.DefaultVersion = ExcelVersion.Xlsx;
                var swb = srcApp.Workbooks.Open(srcPath);

                // 源表按名字索引(忽略历史路径差异)
                var srcSheets = new Dictionary<string, IWorksheet>(StringComparer.OrdinalIgnoreCase);
                foreach (IWorksheet sh in swb.Worksheets)
                    srcSheets[sh.Name] = sh;

                foreach (var en in entries)
                {
                    IWorksheet srcSh;
                    if (!srcSheets.TryGetValue(en.src_sheet, out srcSh))
                    {
                        res.Skipped++;
                        continue;
                    }
                    IWorksheet dstSh = wb.Worksheets[en.dst_sheet];
                    if (dstSh == null)
                    {
                        res.Skipped++;
                        continue;
                    }
                    CopyValue(srcSh.Range[en.src_cell], dstSh.Range[en.dst_cell]);
                    res.Injected++;
                }
            }
            return res;
        }

        static void CopyValue(IRange from, IRange to)
        {
            // 保持类型:数字/文本/日期/布尔
            if (from.HasBoolean)
                to.Boolean = from.Boolean;
            else if (from.HasNumber)
                to.Number = from.Number;
            else if (from.HasDateTime)
                to.DateTime = from.DateTime;
            else
                to.Text = string.IsNullOrEmpty(from.Text) ? null : from.Text;
        }
    }
}
