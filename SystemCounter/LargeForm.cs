using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ChartDirector;

namespace SystemCounter
{
    public partial class LargeForm : Form
    {
        string[] labels = { "T_probe ANT1", "T_probe ANT2", "T_probe ANT3", "HDMI Connector", "RF Splitter" };
        int[] colors    = { 0x8ecae6, 0x219ebc, 0x023047, 0xf77f00, 0xd62828 };
        
        string title = "";
        DataTable myTable = new DataTable();

        int[] color_pallete2 = {
            0xFF80DF,
            0xFF00BF,
            0x804070,
            0x800060,
            0x405080,
            0x002080,
            0xBFFF80,
            0x80FF00,
            0x608040,
            0x408000,
            0xFFDF80,
            0xFFBF00,
            0x807040,
            0x806000,
            0xFF8080,
            0xFF0000,
        };
        

        public LargeForm()
        {
            InitializeComponent();
        }

        public void setChartData(string _title, DataTable _table)
        {
            title = _title;

            myTable = _table;

            myTable.DefaultView.Sort = "Test Count";
            myTable = myTable.DefaultView.ToTable();

        }

        // ===================================================================== //
        //  Draw Chart
        // ===================================================================== //
        private void createChart(WinChartViewer viewer, string chart_title, DataTable table)
        {
            // Prepare DataSet:
            int nSamples = table.Rows.Count;
            double[] chartData = new double[nSamples];
            string[] chartLabel = new string[nSamples];

            for (int i = 0; i < nSamples; i++)
            {
                string val = table.Rows[i]["Test Count"].ToString();
                string testItem = table.Rows[i]["Items"].ToString();
                chartData[i] = Convert.ToDouble(val);
                chartLabel[i] = testItem;
            }

            /*
            double[] top5_data = new double[5];
            string[] top5_label = new string[5];

            for (int i = 0; i < 5; i++)
            {
                int startIndex = nSamples - 5;
                top5_data[i] = chartData[startIndex + i];
                top5_label[i] = chartLabel[startIndex + i];
            }
            */

            // Get chart viewer size:
            int height = viewer.Size.Height;
            int width = viewer.Size.Width;

            // Create chart
            XYChart c = new XYChart(width, height);
            c.setColor(Chart.TextColor, 0xffffff);
            c.setBackground(0x000000);

            ChartDirector.TextBox title = c.addTitle(chart_title, "Arial Bold", 14);
            title.setMargin2(0, 0, 6, 6);
            //c.addLine(20, title.getHeight(), c.getWidth() - 21, title.getHeight(), 0xffffff);

            c.setPlotArea(70, 80, 480, 240, -1, -1, Chart.Transparent, Chart.Transparent);
            c.swapXY();




            // Color pallete:
            int[] colors = color_pallete2;
            Array.Reverse(colors);

            //Console.WriteLine(String.Join(",", color_pallete2));
           // Console.WriteLine(String.Join(",", colors));
            //Array.Reverse(colors);

            int nItems = chartData.Length;
            int[] barColor = new int[nItems];
            for(int i = 0; i < nItems; i++)
            {
                barColor[i] = colors[i];
            }
            /*
            barColor[nItems - 1] = 0xd62828;
            barColor[nItems - 2] = 0xf77f00;
            barColor[nItems - 3] = 0x023047;
            barColor[nItems - 4] = 0x219ebc;
            barColor[nItems - 5] = 0x8ecae6;
            */
            //c.addBarLayer3(data, colors).setBorderColor(Chart.Transparent);
            BarLayer layer = c.addBarLayer3(chartData);
            layer.setBorderColor(Chart.Transparent);
            layer.setAggregateLabelFormat("{value}");

            c.xAxis().setLabels(chartLabel);

            ChartDirector.TextBox textbox = c.xAxis().setLabels(chartLabel);
            textbox.setFontStyle("Arial Bold Italic");
            textbox.setFontSize(10);
            textbox.setFontColor(0xffffff);
            //c.syncYAxis();

            //c.yAxis().setTitle("USD (millions)", "Arial Bold", 10);
            c.yAxis().setColors(Chart.Transparent, Chart.Transparent);
            c.yAxis2().setColors(Chart.Transparent);
            c.xAxis().setTickColor(Chart.Transparent);

            c.xAxis().setLabelStyle("Arial Bold", 8);
            //c.yAxis().setLabelStyle("Arial Bold", 8);
            //c.yAxis2().setLabelStyle("Arial Bold", 8);

            c.packPlotArea(30, title.getHeight() + 25, c.getWidth() - 50, c.getHeight() - 25);

            viewer.Chart = c;
            viewer.ImageMap = c.getHTMLImageMap("clickable", "", "title='{value}'");
            
        }

        private void LargeForm_Load(object sender, EventArgs e)
        {
            
            createChart(winChartViewer1, title, myTable);

            DataTable newtable = myTable;
            newtable.DefaultView.Sort = "Test Count DESC";
            newtable = newtable.DefaultView.ToTable();

            //dataGridView1.DataSource = myTable;
            //dataGridView1.ForeColor = Color.Black;
            //dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dataGridView1.Columns[1].Width = 150;
            
        }

        private void winChartViewer1_SizeChanged(object sender, EventArgs e)
        {
            createChart(winChartViewer1, title, myTable);
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("Are you sure to RESET the test item?", "Test Item Reset", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if(res == DialogResult.OK)
            {
                //int curRow = dataGridView1.CurrentRow.Index;
                //dataGridView1.Rows[curRow].Cells[1].Value = 0;
            }
        }
    }
}
