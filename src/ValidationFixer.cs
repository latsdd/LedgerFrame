using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Syncfusion.XlsIO;

namespace LedgerFrame
{
    // WinForms Spreadsheet 不会对验证公式 INDIRECT(单元格) 即时求值,
    // 导致动态下拉列表不显示。这里在加载后把 INDIRECT 解析成静态区域引用,
    // 并在源单元格被编辑后重新解析,保持原文件的联动行为。
    static class ValidationFixer
    {
        public class Entry
        {
            public IWorksheet Sheet;
            public string SourceCell;
            public IRange Range;
        }

        static readonly Regex IndirectRe =
            new Regex(@"^INDIRECT\(\$?([A-Za-z]+)\$?(\d+)\)$", RegexOptions.IgnoreCase);

        public static List<Entry> ResolveAll(IWorkbook wb)
        {
            var list = new List<Entry>();
            foreach (IWorksheet sh in wb.Worksheets)
            {
                IRange used = sh.UsedRange;
                if (used == null) continue;
                foreach (IRange cell in used.Cells)
                {
                    if (!cell.HasDataValidation) continue;
                    var dv = cell.DataValidation;
                    var m = IndirectRe.Match(dv.FirstFormula ?? "");
                    if (!m.Success) continue;
                    var e = new Entry
                    {
                        Sheet = sh,
                        SourceCell = m.Groups[1].Value.ToUpperInvariant() + m.Groups[2].Value,
                        Range = cell,
                    };
                    ResolveOne(wb, e);
                    list.Add(e);
                }
            }
            return list;
        }

        static void ResolveOne(IWorkbook wb, Entry e)
        {
            string key = e.Sheet.Range[e.SourceCell].Text;
            if (key != null) key = key.Trim();
            if (string.IsNullOrEmpty(key)) return;
            IName name = wb.Names[key];
            if (name == null || name.RefersToRange == null) return;
            e.Range.DataValidation.FirstFormula = name.RefersToRange.AddressGlobal;
        }

        public static void OnCellEdited(IWorkbook wb, List<Entry> entries,
            IWorksheet sheet, int row, int col)
        {
            string addr = ColName(col) + row;
            foreach (var e in entries)
                if (ReferenceEquals(e.Sheet, sheet) && e.SourceCell == addr)
                    ResolveOne(wb, e);
        }

        static string ColName(int col)
        {
            string s = "";
            while (col > 0)
            {
                int r = (col - 1) % 26;
                s = (char)('A' + r) + s;
                col = (col - 1) / 26;
            }
            return s;
        }
    }
}
