using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SystemCounter
{
    public partial class AddDeviceForm : Form
    {
        public AddDeviceForm()
        {
            InitializeComponent();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string device = comboBox1.Text;
            int index = comboBox1.SelectedIndex;
            string id = textBox1.Text;
            DataForm.instance.idTable.Rows.Add(device, id);
            DataForm.instance.dataTable[index].Columns.Add("Items");
            DataForm.instance.dataTable[index].Columns.Add("Test Count");
            DialogResult = DialogResult.OK;
            this.Close();
        }

        private void AddDeviceForm_Load(object sender, EventArgs e)
        {

        }
    }
}
