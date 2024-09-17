using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Qrample
{
    public partial class UserControl2 : UserControl
    {
        public bool CheckBox1State { get { return checkBox1.Checked; } }
        public bool CheckBox2State { get { return checkBox2.Checked; } }
        public string TextBox1Text {  get { return textBox1.Text; } }
        public string comboBox1Selected { get { return comboBox1.SelectedItem.ToString(); } }

        public UserControl2()
        {
            InitializeComponent();
            comboBox1_Init();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var qrHelper = new QRCodesCreatorHelper();
            qrHelper.GenerateQRCodeForSelectedRange();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var BarHelper = new BarCodesCreatorHelper();
            BarHelper.GenerateBarCodeForSelectedRange();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                checkBox2.Checked = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                checkBox1.Checked = false;
            }
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
                MessageBox.Show("Please enter numbers only", 
                    "Validation error", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Warning);
            }
        }
        private void comboBox1_Init()
        {
            var listOfItems = Enum.GetNames(typeof(BarcodeStandard.Type));
            foreach (var item in listOfItems) {
                comboBox1.Items.Add(item);
            }
            comboBox1.SelectedIndex = 0;
        }
    }
}
