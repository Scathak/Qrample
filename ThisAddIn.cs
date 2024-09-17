using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using System.Windows.Forms;

namespace Qrample
{
    public partial class ThisAddIn
    {
        public UserControl1 myUserControl1;
        public UserControl2 myUserControl2;
        public CustomTaskPane CameraTaskPane1;
        public CustomTaskPane CodesTaskPane2;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CameraControlPanel_Init();
            QRCodesControlPanel_Init();
        }
        private void CameraControlPanel_Init()
        {
            myUserControl1 = new UserControl1();
            CameraTaskPane1 = this.CustomTaskPanes.Add(myUserControl1, "Camera panel");
            CameraTaskPane1.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            CameraTaskPane1.Width = 370;
            CameraTaskPane1.Visible = false;

        }
        private void QRCodesControlPanel_Init()
        {
            myUserControl2 = new UserControl2();
            CodesTaskPane2 = this.CustomTaskPanes.Add(myUserControl2, "Code panel");
            CodesTaskPane2.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            CodesTaskPane2.Width = 326;
            CodesTaskPane2.Visible = false;
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
