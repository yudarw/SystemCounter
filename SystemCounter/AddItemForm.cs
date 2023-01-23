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
    public partial class AddItemForm : Form
    {
        int selectionIndex = 0;
        public AddItemForm(int index)
        {
            InitializeComponent();
            selectionIndex = index;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string item = textBox1.Text;
            string count = textBox2.Text;

            DataForm.instance.dataTable[selectionIndex].Rows.Add(item, count);
            DataForm.instance.dataGridView1.DataSource = DataForm.instance.dataTable[selectionIndex];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
