using Microsoft.Office.Tools;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _pane;
        private PdfPaneControl _control;
        private Outlook.Explorer _explorer;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _control = new PdfPaneControl(this.Application);
            _pane = this.CustomTaskPanes.Add(_control, "PDF Embedded Viewer");
            _pane.Width = 520;
            _pane.Visible = true;

            try
            {
                _explorer = this.Application.ActiveExplorer();
                if (_explorer != null)
                    _explorer.SelectionChange += Explorer_SelectionChange;
            }
            catch { }

            // initial refresh
            try { _control.TriggerContextRefresh(); } catch { }
        }

        private void Explorer_SelectionChange()
        {
            try { _control.TriggerContextRefresh(); } catch { }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (_explorer != null)
                    _explorer.SelectionChange -= Explorer_SelectionChange;
            }
            catch { }
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}