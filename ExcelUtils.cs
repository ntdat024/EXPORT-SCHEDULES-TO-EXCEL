using Autodesk.Revit.DB;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace OPENSOURCE
{
    public static class ExcelUtils
    {
        public static string GetFileName(string folderPath, string currentName)
        {
            char[] array = new char[] { '!', '@', '#', '%', '$', '^', '*', '/' };
            string newName = string.Empty;
            foreach (char c in currentName)
            {
                if (!array.Contains(c))
                {
                    newName += c;
                }
                else newName += "_";
            }

            return $"{folderPath}\\{newName}.xlsx";
        }
        public static string GetColumnTitle(int colIndex)
        {
            int index = colIndex;
            string title = string.Empty;
            while (index > 0)
            {
                int num = (index - 1) % 26;
                title = (char)(65 + num) + title;
                index = (index - num) / 26;
            }
            return title;
        }

        public static void BorderTitle(Microsoft.Office.Interop.Excel.Application excel, Worksheet sheet)
        {
            int col = sheet.UsedRange.Columns.Count;
            string letter1 = GetColumnTitle(col) + "1";
            Range range = excel.get_Range("A1", letter1);
            range.Borders.get_Item(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous;
            range.Borders.get_Item(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous;
            range.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous;
            range.Borders.get_Item(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous;
        }

        public static void FormatTitle(Microsoft.Office.Interop.Excel.Application excel, Worksheet sheet, string title)
        {
            string letter1 = GetColumnTitle(sheet.UsedRange.Columns.Count) + "1";
            Range range = excel.get_Range("A1", letter1);
            range.Merge();
            range.RowHeight = 30;
            range.Font.Bold = true;
            sheet.Cells[1, 1] = title;
            sheet.Cells[1, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            sheet.Cells[1, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
            excel.StandardFont = "Tahoma";
            sheet.Columns.AutoFit();
            int col = sheet.UsedRange.Columns.Count;
            double w = 0;
            for (int i = 0; i < col; i++)
            {
                double x = sheet.Columns[i + 1].ColumnWidth;
                w += x;
            }
            if (w <= title.Length)
            {
                double val = (title.Length - w + 5) / col;
                for (int i = 0; i < col; i++)
                {
                    double x = sheet.Columns[i + 1].ColumnWidth;
                    sheet.Columns[i + 1].ColumnWidth = x + val;
                }
            }
            BorderTitle(excel, sheet);
        }

        public static void AddBorder(Microsoft.Office.Interop.Excel.Application excel, Worksheet sheet, ViewSchedule vs)
        {
            Range range = sheet.UsedRange;
            range.Borders.get_Item(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlDot;
            range.Borders.get_Item(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlDot;
            range.Borders.get_Item(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous;
            range.Borders.get_Item(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous;
            range.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous;
            range.Borders.get_Item(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous;

            if (vs.Definition.ShowGrandTotal)
            {
                int row = range.Rows.Count;
                int col = range.Columns.Count;
                string letter1 = "A" + row.ToString();
                string letter2 = GetColumnTitle(col) + row.ToString();
                Range range1 = excel.get_Range(letter1, letter2);
                range1.Borders.get_Item(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlLineStyleNone;
                range1.Borders.get_Item(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlLineStyleNone;
                range1.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous;
                range1.Borders.get_Item(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlLineStyleNone;
                range1.Borders.get_Item(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlLineStyleNone;
            }

        }

        public static bool CheckForContent(Range range)
        {
            range.Select();
            range.Copy();

            string text = Clipboard.GetText().Trim();

            if (string.IsNullOrEmpty(text))
            {
                return true;
            }
            return false;
        }

        public static void BlankRowData(Microsoft.Office.Interop.Excel.Application excel, Worksheet sheet)
        {
            Range usedRange = sheet.UsedRange;
            int row = usedRange.Rows.Count;
            int col = usedRange.Columns.Count;

            for (int i = 1; i <= row; i++)
            {
                string cell1 = "A" + i.ToString();
                string cell2 = GetColumnTitle(col) + i.ToString();
                Range range_i = excel.get_Range(cell1, cell2);
                bool check1 = CheckForContent(range_i);

                if (check1)
                {
                    range_i.Borders.get_Item(XlBordersIndex.xlInsideVertical).LineStyle = XlLineStyle.xlLineStyleNone;
                    range_i.Borders.get_Item(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlLineStyleNone;

                    string cell11 = "A" + (i - 1).ToString();
                    string cell22 = GetColumnTitle(col) + (i - 1).ToString();
                    Range range_b = excel.get_Range(cell11, cell22);
                    bool check2 = CheckForContent(range_b);
                    if (check2 == false)
                    {
                        range_b.Borders.get_Item(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous;
                    }


                    string cell111 = "A" + (i + 1).ToString();
                    string cell222 = GetColumnTitle(col) + (i + 1).ToString();
                    Range range_t = excel.get_Range(cell111, cell222);
                    bool check3 = CheckForContent(range_b);
                    if (check3 == false)
                    {
                        range_t.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous;
                    }

                }
            }

        }

    }
}
