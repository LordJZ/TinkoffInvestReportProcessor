using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;

namespace TinkoffInvestReportProcessor
{
    class Program
    {
        static Program()
        {
            string licenseKey = Environment.GetEnvironmentVariable("SyncfusionLicenseKey");
            if (!string.IsNullOrEmpty(licenseKey))
                Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(licenseKey);
        }

        static readonly ExcelEngine _excel = new ExcelEngine();

        static void Main(string[] args)
        {
            CultureInfo.CurrentCulture =
                CultureInfo.CurrentUICulture =
                    CultureInfo.DefaultThreadCurrentCulture =
                        CultureInfo.DefaultThreadCurrentUICulture = ruRU;

            //Environment.CurrentDirectory = "D:\\Dropbox\\Alexander Zhuravlev Tax Reports\\";

            Directory.CreateDirectory("fixed");

            string[] files = Directory.GetFiles(".", "broker-report-*.xlsx", SearchOption.AllDirectories).OrderBy(x => x).ToArray();

            foreach (string file in files.Select(x =>
            {
                if (x.StartsWith(".\\"))
                    x = x.Substring(2);

                return x;
            }))
            {
                if (file.EndsWith("-fixed.xlsx"))
                    continue;

                string outputFilename = Path.GetFileNameWithoutExtension(file);
                string subfolder = Path.GetDirectoryName(file);
                if (!string.IsNullOrEmpty(subfolder))
                    outputFilename += "-" + subfolder.Replace('\\', '/').Replace('/', '-');
                outputFilename += "-fixed.xlsx";
                string outputFilepath = Path.Combine("fixed", outputFilename);

                if (File.Exists(outputFilepath))
                    continue;

                IWorkbook wb = _excel.Excel.Workbooks.Create(1);
                IWorksheet ws = wb.Worksheets[0];

                bool autowidth = false;
                int row = 1;
                foreach ((string text, DataTable dt) in ReadDataTables(file))
                {
                    if (dt != null)
                    {
                        ++row;
                        ws[row, 1].CellStyle.Font.Bold = true;
                        ws[row++, 1].Value = dt.TableName;

                        int rowsImported = ws.ImportDataTable(dt, true, row, 2);

                        IRange table = ws[row, 2, row + rowsImported, 2 - 1 + dt.Columns.Count];

                        // fixup time formatted as date
                        if (table.Row != table.LastRow)
                        {
                            foreach (DataColumn col in dt.Columns)
                            {
                                if (col.ColumnName.Contains("время", StringComparison.CurrentCultureIgnoreCase) &&
                                    !col.ColumnName.Contains("дата", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    IRange timeRange = table[table.Row + 1, table.Column + col.Ordinal,
                                                             table.LastRow, table.Column + col.Ordinal];

                                    foreach (IRange cell in timeRange)
                                    {
                                        if (cell.Value2 is DateTime dateTime)
                                            cell.TimeSpan = dateTime.TimeOfDay;
                                    }
                                    timeRange.NumberFormat = "[$-F400]h:mm:ss\\ AM/PM";
                                }
                            }
                        }

                        if (!autowidth)
                        {
                            for (int i = table.Column, l = table.LastColumn; i <= l; i++)
                                ws.AutofitColumn(i);

                            autowidth = true;
                        }

                        ws.ListObjects.Create("Table" + row, table);
                        row += rowsImported;
                        row += 2;
                    }
                    else
                    {
                        ws[row, 1].Value = text;
                        ++row;
                    }
                }

                using (FileStream fs = File.Create(outputFilepath))
                    wb.SaveAs(fs);
            }


        }

        static IEnumerable<(string, DataTable)> ReadDataTables(string file)
        {
            using FileStream fs = File.OpenRead(file);
            IWorkbook wb = _excel.Excel.Workbooks.Open(fs);
            try
            {
                foreach (var t in ReadDataTables(wb.Worksheets[0]))
                    yield return t;
            }
            finally
            {
                wb.Close();
            }
        }

        static IEnumerable<(string, DataTable)> ReadDataTables(IWorksheet ws)
        {
            int row = 1;
            while (true)
            {
                DataTable dt;
                string[] texts;
                (dt, row, texts) = ReadDataTable(ws, row);
                foreach (string t in texts)
                    yield return (t, null);

                if (dt == null)
                    yield break;

                yield return (null, dt);
            }
        }

        static (DataTable, int, string[]) ReadDataTable(IWorksheet ws, int row)
        {
            List<string> texts = new List<string>();

            int lastRow = ws.UsedRange.LastRow;
            int lastColumn = ws.UsedRange.LastColumn;

            bool IsHeaderCell(int r, int col) => (ws[r, col].MergeArea?.Count ?? 1) < 60 && ws[r, col].Value?.Length > 0;

            int col = (from r in Enumerable.Range(1, 10)
                       from c in Enumerable.Range(1, 10)
                       select ws[r, c]).First(cell => cell.Text?.Length > 0).Column;

            bool IsHeader(int r) => IsHeaderCell(r, col) || col > 1 && IsHeaderCell(r, col - 1);

            bool IsPageBreak(int r)
            {
                string s = "";
                for (int c = col; c <= lastColumn && s.Length < 15; ++c)
                {
                    s += ws[r, c].Value;
                }

                return s.Length < 15 && pageBreak.IsMatch(s);
            }

            bool IsTableName() => ws[row, col].MergeArea?.Count > 100 &&
                                  (IsHeader(row + 1) || IsPageBreak(row + 1) && IsHeader(row + 2));

            while (!IsTableName())
            {
                texts.Add(ws[row, col].Value);

                if (row > lastRow)
                    return (default, default, texts.ToArray());

                ++row;
            }

            DataTable dt = new DataTable(ws[row, col].Value);
            ++row;

            if (IsPageBreak(row))
                ++row;

            (int, string)[] headers =
                Enumerable.Range(col, lastColumn).Select(x => (x, ws[row, x]))
                          .Where(t => t.Item2.Value != "")
                          .Select(t => (t.x, t.Item2.Value.Replace("\n", "")))
                          .ToArray();

            foreach ((_, string val) in headers)
                dt.Columns.Add(val);

            Dictionary<string, object> dict = new Dictionary<string, object>();
            while (++row < lastRow && !IsTableName())
            {
                dict.Clear();

                bool isSecondaryHeader = headers.All(h => ws[row, h.Item1].Value.Replace("\n", "") == h.Item2);
                if (isSecondaryHeader)
                    continue;

                foreach ((int colOrdinal, string colName) in headers)
                {
                    object val = ws[row, colOrdinal].Value2;
                    if (val != null && !"".Equals(val))
                    {
                        dict.Add(colName, val);
                    }
                }

                if (dict.Count > 0)
                {
                    DataRow dr = dt.NewRow();

                    foreach ((string colOrdinal, object val) in dict)
                        dr[colOrdinal] = val;

                    dt.Rows.Add(dr);
                }
            }

            return (dt, row, texts.ToArray());
        }

        static Regex pageBreak = new Regex(@"\d+ из \d+");
        static CultureInfo ruRU = CultureInfo.GetCultureInfo("ru-RU");
    }
}
