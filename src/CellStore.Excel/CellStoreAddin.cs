using ExcelDna.Integration;
using CellStore.Excel.Tools;

namespace CellStore.Excel.Api
{
    
    /// <summary>
    /// The CellStore AddIn
    /// </summary>
    public class CellStoreAddin : IExcelAddIn
    {

        public void AutoOpen()
        {
            ExcelAsyncUtil.Initialize();
            Utils.initLogWriter();
            //ExcelIntegration.RegisterUnhandledExceptionHandler(
            //    ex => "!!! EXCEPTION: " + ex.ToString());
        }

        public void AutoClose()
        {
            Utils.closeLogWriter();
        }
    }

}
