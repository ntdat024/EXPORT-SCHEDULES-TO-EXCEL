using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CommunityToolkit.Mvvm.Input;
using PropertyChanged;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace OPENSOURCE
{
    [AddINotifyPropertyChangedInterface]
    public class ExportScheduleToExcelVM
    {
        public Document Doc { get; set; }
        public ExportScheduleToExcelView MainWindow { get; set; }
        public ObservableCollection<ViewSchedule> ListSchedules { get; set; }
        public List<ViewSchedule> SelectedSchedules { get; set; }
        public string TextFilter { get; set; }
        public string FolderPath { get; set; }
        public bool ExportToOneFile { get; set; }

        public ICommand OKCmd { get; }
        public ICommand CancelCmd { get; }
        public ICommand BrowseCmd { get; }
        public ICommand TextFilterChanged { get; }

        public ExportScheduleToExcelVM (UIApplication uiApp)
        {
            Doc = uiApp.ActiveUIDocument.Document;
            SelectedSchedules = new List<ViewSchedule>();
            ExportToOneFile = true;

            OKCmd = new RelayCommand<Window>((p) => { Ok(); });
            CancelCmd = new RelayCommand<Window>((p) => { MainWindow.Close(); });
            BrowseCmd = new RelayCommand<Window>((p) => { Browse(); });
            TextFilterChanged = new RelayCommand<Window>((p) => { UpdateListSchedules(); });
            
        }

        private void UpdateListSchedules()
        {
            var filterList = ExportScheduleToExcelUtils.GetViewSchedules(Doc, TextFilter);
            ListSchedules = new ObservableCollection<ViewSchedule>(filterList);
        }
        private void Browse()
        {
            FolderPath = ExportScheduleToExcelUtils.GetFolderPath();
        }
        private void Ok ()
        {
            var lbx = MainWindow.FindName("lbx_Schedules") as ListBox;
            if (lbx == null)
            {
                MessageBox.Show("Can not get data!", "Message");
                MainWindow.Close();
                return;
            }

            if (string.IsNullOrEmpty(FolderPath))
            {
                MessageBox.Show("Select a folder to save file!", "Message");
            }
            else
            {
                foreach (var item in lbx.SelectedItems)
                {
                    SelectedSchedules.Add(item as ViewSchedule);
                }

                if (!SelectedSchedules.Any())
                {
                    MessageBox.Show("0 scheduled selected!", "Message");
                }
                else
                {
                    MainWindow.Close();
                    ExportScheduleToExcelUtils.ExportSchedules(this);
                }
            }

        }

        
    }
}
