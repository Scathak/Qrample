using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Qrample
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CameraTaskPane1.Visible = toggleButton1.Checked;
            //if (!toggleButton1.Checked) { Globals.ThisAddIn.myUserControl.videoSource.Stop(); }
            toggleButton1.Label = toggleButton1.Checked ? "CamPane Off" : "CamPane On";
        }

        private void toggleButton2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CodesTaskPane2.Visible = toggleButton2.Checked;
            toggleButton2.Label = toggleButton2.Checked ? "CodePane Off" : "CodePane On";
        }
    }
}
