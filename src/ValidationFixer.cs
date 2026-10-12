using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Syncfusion.XlsIO;

namespace LedgerFrame
{
    // 联动下拉重建:原文件用 INDIRECT(源单元格) 做级联验证(如 B2 选街道 ->
    // B3/D3 等列出该街道的村)。模型生成期 INDIRECT 被冻结成静态区域,
    // 这里按 validation_map.json 在启动时和源单元格被编辑后重建下拉源。
    static class ValidationFixer
    {
        public class MapEntry
        {
            public string sheet { get; set; }
            public List<string> cells { get; set; }
            public string source { get; set; }
        }

        public class Entry
        {
            public IWorksheet Sheet;
            public string SourceCell;
            public List<IRange> Ranges = new List<IRange>();
        }

        public static List<Entry> LoadMap(IWorkbook wb, string mapJson)
        {
            return LoadMapForSheet(wb, mapJson, null);
        }

        public static List<Entry> LoadMapForSheet(IWorkbook wb, string mapJson, string onlySheet)
        {
            var list = new List<Entry>();
            var maps = JsonSerializer.Deserialize<List<MapEntry>>(mapJson);
            foreach (var m in maps)
            {
                if (onlySheet != null && m.sheet != onlySheet) continue;
                IWorksheet sh;
                try { sh = wb.Worksheets[m.sheet]; }
                catch { continue; }
                if (sh == null) continue;
                var e = new Entry { Sheet = sh, SourceCell = m.source };
                foreach (var c in m.cells)
                {
                    try
                    {
                        var r = sh.Range[c];
                        if (r.HasDataValidation) e.Ranges.Add(r);
                    }
                    catch { }
                }
                if (e.Ranges.Count > 0)
                {
                    ResolveOne(wb, e);
                    list.Add(e);
                }
            }
            return list;
        }

        static void ResolveOne(IWorkbook wb, Entry e)
        {
            string key = null;
            try { key = e.Sheet.Range[e.SourceCell].Text; } catch { return; }
            if (key != null) key = key.Trim();
            if (string.IsNullOrEmpty(key)) return;
            IName name = null;
            try { name = wb.Names[key]; } catch { }
            if (name == null || name.RefersToRange == null) return;
            string target = name.RefersToRange.AddressGlobal;
            foreach (var r in e.Ranges)
            {
                try { r.DataValidation.FirstFormula = target; } catch { }
            }
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
