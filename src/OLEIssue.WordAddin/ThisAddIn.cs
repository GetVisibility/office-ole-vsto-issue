using Microsoft.Office.Interop.Word;
using OLEIssue.Common;

namespace OLEIssue.WordAddin
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

            ((ApplicationEvents2_Event)Application).NewDocument += Application_OnNewDocument;
            Application.DocumentOpen += Application_DocumentOpen;
            Application.DocumentBeforeSave += Application_DocumentBeforeSave;
            Application.DocumentBeforePrint += Application_BeforePrint;
            Application.DocumentBeforeClose += Application_DocumentBeforeClose;

            Application.WindowActivate += Application_WindowActivate;
            Application.WindowDeactivate += Application_WindowDeactivate;
        }

        private void Application_OnNewDocument(Document doc)
        {
            try
            {
                Logger.Log().Debug("Document OnNewDocument");
            }
            finally
            {
                OfficeUtils.ConditionalReleaseComObject(doc);
            }
        }

        private void Application_DocumentOpen(Document doc)
        {
            try
            {
                Logger.Log().Debug("Document DocumentOpen");
            }
            finally
            {
                OfficeUtils.ConditionalReleaseComObject(doc);
            }
        }

        private void Application_DocumentBeforeSave(Document doc, ref bool saveAsUiRef, ref bool cancel)
        {
            try
            {
                if (cancel) return;
                Logger.Log().Debug("Document BeforeSave");
            }
            finally
            {
                OfficeUtils.ConditionalReleaseComObject(doc);
            }
        }

        private void Application_DocumentBeforeClose(Document doc, ref bool cancel)
        {
            try
            {
                if (cancel) return;
                Logger.Log().Debug("Document BeforeClose");
            }
            finally
            {
                OfficeUtils.ConditionalReleaseComObject(doc);
            }
        }

        private void Application_BeforePrint(Document doc, ref bool cancel)
        {
            try
            {
                if (cancel) return;
                Logger.Log().Debug("Document BeforePrint");
            }
            finally
            {
                OfficeUtils.ConditionalReleaseComObject(doc);
            }
        }

        private void Application_WindowActivate(Document doc, Window wn)
        {
            try
            {
                Logger.Log().Debug("Document WindowActivate");
            }
            finally
            {
                OfficeUtils.ConditionalReleaseComObject(doc);
                OfficeUtils.ConditionalReleaseComObject(wn);
            }
        }

        private void Application_WindowDeactivate(Document doc, Window wn)
        {
            try
            {
                Logger.Log().Debug("Document WindowDeactivate");
            }
            finally
            {
                OfficeUtils.ConditionalReleaseComObject(doc);
                OfficeUtils.ConditionalReleaseComObject(wn);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Logger.Log().Debug("Application start shutdown");

            ((ApplicationEvents2_Event)Application).NewDocument -= Application_OnNewDocument;
            Application.DocumentOpen -= Application_DocumentOpen;
            Application.DocumentBeforeSave -= Application_DocumentBeforeSave;
            Application.DocumentBeforePrint -= Application_BeforePrint;
            Application.DocumentBeforeClose -= Application_DocumentBeforeClose;

            Application.WindowActivate -= Application_WindowActivate;
            Application.WindowDeactivate -= Application_WindowDeactivate;

            Logger.Log().Debug("Application shutdown done");
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
