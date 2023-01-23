using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Bytescout.Spreadsheet;

namespace SystemCounter
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }


        // ===================================================================== //
        // Read Configuration File 
        // ===================================================================== //
        private Spreadsheet readExcel(string filePath)
        {
            Spreadsheet document = new Spreadsheet();
            try
            {
                document.LoadFromFile(filePath);
                DataTable myTable = document.ExportToDataTable(0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ReadCSV: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return document;
        }


        // ===================================================================== //
        //  Form 2 Load
        // ===================================================================== //
        private void Form2_Load(object sender, EventArgs e)
        {
           
            DataTable myTable = Form1.instance.dataTable[0];
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.DataSource = myTable;

            dataGridView1.Columns["Testing Time"].Width = 100;
            dataGridView1.Columns["Testing Time"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns["Testing Time"].SortMode = DataGridViewColumnSortMode.Programmatic;
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            /*
            dataGridView1.Rows.Add("1", "T-Probe 1", 100);
            dataGridView1.Rows.Add("2", "T-Probe 2", 99);
            dataGridView1.Rows.Add("3", "T-Probe 3", 1289);
            dataGridView1.Rows.Add("4", "T-Probe 4", 32);
            dataGridView1.Rows.Add("5", "T-Probe 5", 66);
            dataGridView1.Rows.Add("6", "T-Probe 6", 198);
            dataGridView1.Rows.Add("7", "T-Probe 7", 2089);
            dataGridView1.Rows.Add("8", "T-Probe 8", 3012);
            dataGridView1.Rows.Add("9", "T-Probe 9", 43);
            dataGridView1.Rows.Add("10", "T-Probe 10", 1);
            dataGridView1.Rows.Add("11", "T-Probe 11", 6);
            dataGridView1.Rows.Add("12", "T-Probe 12", 178);
            */
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Spreadsheet xlFile = readExcel("device-status-report.xlsx");
            //DataTable myTable = xlFile.ExportToDataTable(0);            
            //dataGridView1.DataSource = myTable;
            //dataGridView1.Rows.Add("1", "T-Probe 1", 100);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Spreadsheet document = new Spreadsheet();
            Worksheet sheet = document.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cell(0, 0).Value = "This is the input from the program";
            
            document.SaveAs("device-status-report.xlsx");
            
        }

    }
}
