using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;

namespace OPENSOURCE
{
    internal class ExportScheduleToExcelUtils
    {
        private static int MAX_LENGTH = 31;
        public static void Main(UIApplication uiApp)
        {
            try
            {
                UIDocument uidoc = uiApp.ActiveUIDocument;
                Document doc = uidoc.Document;

                var collector = GetViewSchedules(doc, string.Empty);

                if (collector.Count == 0)
                {
                    MessageBox.Show("There is no schedules in the current model!", "Message");
                    return;
                }

                var vm = new ExportScheduleToExcelVM(uiApp);
                vm.ListSchedules = new ObservableCollection<ViewSchedule>(collector);

                var window = new ExportScheduleToExcelView() { DataContext = vm };
                vm.MainWindow = window;
                window.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Main");
                return;
            }
            
        }

        public static List<ViewSchedule> GetViewSchedules(Document doc, string filterName)
        {
            var collector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Schedules)
                .WhereElementIsNotElementType()
                .Cast<ViewSchedule>()
                .Where(x => !x.Name.Contains("<"))
                .OrderBy(x => x.Name)
                .ToList();

            if (string.IsNullOrEmpty(filterName))
            {
                return collector;
            }

            var filterList = collector.FindAll(x => x.Name.ToLower().Contains(filterName.ToLower()));
            return filterList;
        }

        private static void WriteContent(Worksheet sheet, ViewSchedule vs)
        {
            TableData table = vs.GetTableData();
            TableSectionData section = table.GetSectionData(SectionType.Body);

            int row = section.NumberOfRows;
            int col = section.NumberOfColumns;

            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    try
                    {
                        int r = i + 2;
                        int c = j + 1;
                        sheet.Cells[r, c] = vs.GetCellText(SectionType.Body, i, j);
                    }
                    catch { }
                }
            }
        }


        private static void ExportToOneFile(ExportScheduleToExcelVM vm)
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
            }
            catch
            {
                MessageBox.Show("Please install excel!", "Message");
                return;
            }

            Workbook wbook = excel.Workbooks.Add();
            Sheets newsheet = wbook.Sheets;

            foreach (ViewSchedule vs in vm.SelectedSchedules)
            {
                Worksheet sheet = (Worksheet)newsheet.Add(newsheet[1], Type.Missing, Type.Missing, Type.Missing);
                excel.ActiveWindow.DisplayGridlines = false;

                WriteContent(sheet, vs);
                ExcelUtils.FormatTitle(excel, sheet, vs.Name);
                ExcelUtils.AddBorder(excel, sheet, vs);
                ExcelUtils.BlankRowData(excel, sheet);

                if (vs.Name.Length <= MAX_LENGTH) sheet.Name = vs.Name;

            }

            string fileName = ExcelUtils.GetFileName(vm.FolderPath, vm.Doc.Title);
            wbook.SaveAs(fileName);
            wbook.Close(true, null, null);
            excel.Quit();

            Marshal.ReleaseComObject(wbook);
            Marshal.ReleaseComObject(excel);

            var question = MessageBox.Show("Exported! Do you want to open folder?", "Message", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (question == MessageBoxResult.Yes)
            {
                Process.Start(vm.FolderPath);
            }
        }

        private static void ExportToMultiFiles(ExportScheduleToExcelVM vm)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
            }
            catch
            {
                MessageBox.Show("Please install excel!", "Message");
                return;
            }

            foreach (ViewSchedule vs in vm.SelectedSchedules)
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Workbook wbook = excel.Workbooks.Add();
                Worksheet sheet = wbook.ActiveSheet;
                excel.ActiveWindow.DisplayGridlines = false;


                WriteContent(sheet, vs);
                ExcelUtils.FormatTitle(excel, sheet, vs.Name);
                ExcelUtils.AddBorder(excel, sheet, vs);
                ExcelUtils.BlankRowData(excel, sheet);

                string fileName = ExcelUtils.GetFileName(vm.FolderPath, vs.Name);
                wbook.SaveAs(fileName);
                wbook.Close(true, null, null);
                excel.Quit();

                Marshal.ReleaseComObject(sheet);
                Marshal.ReleaseComObject(wbook);
                Marshal.ReleaseComObject(excel);

            }

            var question = MessageBox.Show("Exported! Do you want to open folder?", "Message", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (question == MessageBoxResult.Yes)
            {
                Process.Start(vm.FolderPath);
            }
        }

        public static string GetFolderPath()
        {
            var fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                return fbd.SelectedPath;
            }
            return string.Empty;
        }

        public static void ExportSchedules (ExportScheduleToExcelVM vm)
        {
            if (vm.ExportToOneFile)
            {
                ExportToOneFile(vm);
            }
            else ExportToMultiFiles(vm);
        }

    }
}
