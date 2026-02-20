using Microsoft.Office.Interop.Excel;
using OLEIssue.Common;

namespace OLEIssue.ExcelAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (OfficeUtils.IsStartedByAutomation() && !Application.Visible)
            {
                Logger.Log().Warning("Not loading addin because loaded as embedded application.");
                return;
            }

            ((AppEvents_Event)Application).NewWorkbook += Application_NewWorkbook;
            Application.WorkbookOpen += Application_WorkbookOpen;
            Application.WorkbookNewSheet += Application_WorkbookNewSheet;
            Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
            Application.WorkbookAfterSave += Application_WorkbookAfterSave;
            Application.WorkbookBeforePrint += Application_WorkbookBeforePrint;
            Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;

            Application.WindowActivate += Application_WindowActivate;
            Application.WindowDeactivate += Application_WindowDeactivate;
        }

        private void Application_WorkbookBeforePrint(Workbook wb, ref bool cancel)
        {
            Logger.Log().Debug("Workbook WorkbookBeforePrint");
        }

        private void Application_WorkbookAfterSave(Workbook wb, bool success)
        {
            Logger.Log().Debug("Workbook WorkbookAfterSave");
        }

        private void Application_WorkbookBeforeSave(Workbook wb, bool saveAsUi, ref bool cancel)
        {
            Logger.Log().Debug("Workbook WorkbookBeforeSave");
        }

        private void Application_WorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            if (cancel)
            {
                return;
            }

            Logger.Log().Debug("Workbook BeforeClose");
        }

        private void Application_WorkbookNewSheet(Workbook wb, object sh)
        {
            Logger.Log().Debug("Workbook WorkbookNewSheet");
        }

        private void Application_WorkbookOpen(Workbook wb)
        {
            Logger.Log().Debug("Workbook WorkbookOpen");
        }

        private void Application_NewWorkbook(Workbook wb)
        {
            Logger.Log().Debug("Workbook NewWorkbook");
        }

        private void Application_WindowActivate(Workbook wb, Window wn)
        {
            Logger.Log().Debug("Workbook WindowActivate");
        }

        private void Application_WindowDeactivate(Workbook wb, Window wn)
        {
            Logger.Log().Debug("Workbook WindowDeactivate");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
