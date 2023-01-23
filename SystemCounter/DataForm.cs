using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using System.IO;


namespace SystemCounter
{
    public partial class DataForm : Form
    {
        public static DataForm instance;
        public DataGridView deviceTable;

        DataTableCollection tableCollection;
        TableReader tableReader = new TableReader();

        public DataForm()
        {
            InitializeComponent();
            instance = this;
            deviceTable = dataGridView2;
        }

        public DataTable[] dataTable = new DataTable[16];
        public DataTable idTable = new DataTable();

        public void printDataTable(DataTable table)
        {
            try
            {
                Console.WriteLine("========================== Print Data Table ===========================");
                Console.WriteLine("Size : {0}x{1}\n", table.Rows.Count, table.Columns.Count);
                int row = table.Rows.Count;
                int col = table.Columns.Count;

                foreach (DataRow data in table.Rows)
                {
                    for (int i = 0; i < col; i++)
                    {
                        Console.Write(data[i] + "\t\t");
                    }
                    Console.Write("\n");
                }
                Console.WriteLine("====================================================================== \n\n");
            }
            catch (Exception ex)
            {
                MessageBox.Show("print_data_table: " + ex.Message);
            }
        }

        /*
        private void table_to_sheet(Workbook book, string sheetName, DataTable table)
        {
            Worksheet sheet = book.Worksheets.Add(sheetName);

            int nRow = table.Rows.Count;
            int nCol = table.Columns.Count;

            sheet.Range[1, 1].Value = table.Columns[0].ColumnName;
            sheet.Range[1, 1].Style.Font.IsBold = true;
            sheet.Range[1, 1].Style.Color = Color.LightYellow;

            sheet.Range[1, 2].Value = table.Columns[1].ColumnName;
            sheet.Range[1, 2].Style.Font.IsBold = true;
            sheet.Range[1, 2].Style.Color = Color.LightYellow;

            for (int i = 0; i < nRow; i++)
                for (int j = 0; j < nCol; j++)
                    sheet.Range[i + 2, j + 1].Value = table.Rows[i][j].ToString();
        }
        */

        private void saveExcelFile()
        {
              
        }


        private void update_table(DataTable table)
        {
            DataTable newtable = table;
            
            newtable.DefaultView.Sort = "Test Count DESC";
            newtable = newtable.DefaultView.ToTable();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.DataSource = newtable;
            dataGridView1.Columns[1].Width = 130;
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;            
        }


        private void load_excel(string path)
        {
            try
            {
                using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        System.Data.DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        tableCollection = result.Tables;
                        
                        idTable = tableCollection["DeviceID"];
                        printDataTable(idTable);


                        /*
                        listBox.Items.Clear();
                              
                        int nID = idTable.Rows.Count;
                        for (int i = 0; i < nID; i++)
                            listBox.Items.Add(idTable.Rows[i][0]);
                        */

                        for (int i = 0; i < 16; i++)
                        {
                            // Device ID:
                            //string id = idTable.Rows[i][0].ToString();
                            //string title = String.Format("Device {0} {1}", i + 1, id);
                            //chartTitles[i] = title;

                            // Status of each device:
                            dataTable[i] = tableCollection[i + 1];
                            printDataTable(dataTable[i]);
                        }
                        update_table(dataTable[0]);
                        reader.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        private void save_excel_file()
        {
            using(var stream = File.OpenWrite("testing.xlsx"))
            {
                using(IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    System.Data.DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    { 
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });
                }
            }
        }

        private void show_chart_onLoad()
        {
        }

        
        private void btnLoad_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel file (*.xlsx)|*.xlsx";
            openFileDialog1.FileName = "TestItem Record Data.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog1.FileName;
                load_excel(path);
            }
        }

        private void listBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*
            int nID = listBox.SelectedIndex;
            if (nID == -1) return;
            textBox1.Text = listBox.SelectedItem.ToString();
            update_table(dataTable[nID]);
            */
            
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Next item is pressed!");
            /*
            int nID = listBox.SelectedIndex;
            nID++;
            if (nID > listBox.Items.Count - 1) return;
            listBox.SelectedIndex = nID;
            Console.WriteLine(nID);
            textBox1.Text = listBox.SelectedItem.ToString();
            update_table(dataTable[nID]);
            */
        }

        private void btnPrev_Click(object sender, EventArgs e)
        {
            /*
            int nID = listBox.SelectedIndex;
            nID--;
            if (nID < 0) return;
            listBox.SelectedIndex = nID;
            Console.WriteLine(nID);
            textBox1.Text = listBox.SelectedItem.ToString();
            update_table(dataTable[nID]);
            */
        }

        public void add_device(string device, string id)
        {
            dataGridView2.Rows.Add(device, id);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            /*
            AddDeviceForm form = new AddDeviceForm();
            if(form.ShowDialog() == DialogResult.OK)
            {
                update_item_table();                
            }*/
            int nDevice = idTable.Rows.Count + 1;
            string device = "Device " + nDevice;
            idTable.Rows.Add(device, 0);
            dataTable[nDevice - 1].Columns.Add("Items");
            dataTable[nDevice - 1].Columns.Add("Test Count",typeof(Int16));
            Form1.instance.update_chart();
        }

        // Update Chart:
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Form1.instance.update_chart();
        }

        private void DataForm_Load(object sender, EventArgs e)
        {
            update_item_table();
        }

        private void update_item_table()
        {
            
            dataGridView2.DataSource = idTable;
            dataGridView2.Columns[0].Width = 70;
            dataGridView2.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            
            
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            //saveExcelFile();
            //tableReader.print_allTable(dataTable);

            tableReader.saveExcel(idTable, dataTable);
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int id = dataGridView2.SelectedRows.Count;
            
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            int id = dataGridView2.CurrentRow.Index;
            textBox1.Text = dataGridView2.Rows[id].Cells[1].Value.ToString();
            dataGridView1.DataSource = dataTable[id];
            dataGridView1.Columns[1].Width = 130;
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            Form1.instance.update_chart();
            //update_table(dataTable[id]);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int id = dataGridView2.CurrentRow.Index;
            string id_str = dataGridView2.Rows[id].Cells[1].Value.ToString();
            DialogResult res = MessageBox.Show("Are you sure want to remove the device " + id_str + "?", "Remove Device", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (res == DialogResult.Yes)
            {
                idTable.Rows[id].Delete();
                Form1.instance.update_chart();
                Form1.instance.enable_chart(id, false);
                //update_item_table();
            }
        }


        // Add Item:
        private void button3_Click(object sender, EventArgs e)
        {
            int currentSelectionIndex = dataGridView2.CurrentRow.Index;            
            dataTable[currentSelectionIndex].Rows.Add("New Item", "0");
            dataTable[currentSelectionIndex].DefaultView.Sort = "Test Count";
            dataTable[currentSelectionIndex] = dataTable[currentSelectionIndex].DefaultView.ToTable();
            dataGridView1.DataSource = dataTable[currentSelectionIndex];
            

            //AddItemForm form = new AddItemForm(currentSelectionIndex);
            //form.ShowDialog();
        }

        private void dataGridView2_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int id = dataGridView2.CurrentRow.Index;
            textBox1.Text = dataGridView2.Rows[id].Cells[1].Value.ToString();
            Form1.instance.update_chart();
        }

        private void DataForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form1.instance.update_chart();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            Form1.instance.update_chart();
        }
    }
}
