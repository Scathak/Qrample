using System.Windows.Forms;
using System;
using AForge.Video.DirectShow;
using AForge.Video;
using System.Drawing;
using System.Threading;

namespace Qrample
{
    public partial class UserControl1 : UserControl
    {
        private int selectedItemNumber = -1;
        public bool checkBox1Checked { get { return checkBox1.Checked; } }
        public bool checkBox2Checked { get { return checkBox2.Checked; } }
        public bool checkBox3Checked { get { return checkBox3.Checked; } }
        public bool checkBox4Checked { get { return checkBox4.Checked; } }
        public bool checkBox5Checked { get { return checkBox5.Checked; } }
        public bool checkBox6Checked { get { return checkBox6.Checked; } }
        public string comboBox1Selected { get { return comboBox1.SelectedItem.ToString(); } }
        public int comboBox1SelectedIndex { get { return comboBox1.SelectedIndex; } }
        public string textBox1Text { get { return textBox1.Text; } }
        public PictureBox pictureBox {get {return pictureBox1;} }

        public AllCodesReader codesReader;

        public UserControl1()
        {
            InitializeComponent();
            codesReader = new AllCodesReader(this);
            codesReader.PopulateCameras(this.comboBox1);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            var newString = textBox1.Text;
            if (!string.IsNullOrEmpty(newString))
            {
                foreach (string element in comboBox1.Items)
                {
                    if (element.Contains(newString)) return;
                }
                comboBox1.Items.Add(comboBox1.Items.Count + ". " + newString);
            }
            return;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            codesReader.InsertDecodedQR("test1");
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedCamera = comboBox1Selected;
            if (!string.IsNullOrEmpty(selectedCamera))
            {
                selectedItemNumber = comboBox1.SelectedIndex;
                codesReader.USBCameraAddress = comboBox1.SelectedIndex;
                codesReader.StopCamera();
                codesReader.startSelectedCamera(selectedCamera);
            }
        }
    }
}
