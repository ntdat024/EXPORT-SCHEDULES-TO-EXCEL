using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace OPENSOURCE
{
    [Transaction(TransactionMode.Manual)]
    public class ExportSchedulesCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;

            ExportScheduleToExcelUtils.Main(uiApp);

            return Result.Succeeded;
        }


    }
}
