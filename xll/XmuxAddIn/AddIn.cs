using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace XmuxAddIn
{
    public sealed class AddIn : IExcelAddIn
    {
        private CustomTaskPane pane;

        public void AutoOpen()
        {
            var control = new PaneControl();
            pane = CustomTaskPaneFactory.CreateCustomTaskPane(control, "Xmux");
            if (pane == null)
            {
                throw new System.InvalidOperationException("Excel could not create the Xmux custom task pane.");
            }

            pane.Width = 420;
            pane.Visible = true;
        }

        public void AutoClose()
        {
            if (pane != null)
            {
                pane.Delete();
                pane = null;
            }
        }
    }
}
