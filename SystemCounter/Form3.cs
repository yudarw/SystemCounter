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
    public partial class Form3 : Form
    {
        public static Form3 instance;

        int maxDevices = 16;
        int nDevice = 0;

        public Form3()
        {
            InitializeComponent();
            instance = this;
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            //dataGridView1.Rows.Add("T-Probe 1", 212);
            //dataGridView1.Rows.Add("T-Probe 2", 1023);
            //dataGridView1.Rows.Add("T-Probe 3", 766);
            //dataGridView1.Rows.Add("T-Probe 4", 890);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            /*
            nDevice++;
            if(nDevice > maxDevices)
            {
                MessageBox.Show("Can not add new device. Already reach the maximum devices.");
                return;
            }
            string id = "3770201";
            string items = String.Format("Device {0} [{1}]", nDevice, id);
            listBox.Items.Add(items);
            */
            AddDeviceForm form = new AddDeviceForm();
            form.ShowDialog();
        }

        private void listBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //string id = listBox.SelectedItem.ToString();
            int id = listBox.SelectedIndex;
            Console.WriteLine(id);
            //Console.WriteLine(id);
            //labelID.Text = listBox.SelectedItem.ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(listBox.SelectedItem == null)
            {
                MessageBox.Show("Can not Add new item! No device is selected!", "Add New Items", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataTable idTable = Form1.instance.idTable;


            string id_str = listBox.SelectedItem.ToString();
            string id = id_str.Trim();
            Console.WriteLine(id);



            //AddItemForm form = new AddItemForm();
            //form.ShowDialog();
        }

        private void btnDeleteItem_Click(object sender, EventArgs e)
        {
            if (listBox.SelectedItem == null)
            {
                MessageBox.Show("Can not Add new item! No device is selected!", "Add New Items", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
    }
}
