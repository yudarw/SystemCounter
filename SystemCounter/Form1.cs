using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;
using ChartDirector;
using ExcelDataReader;
using System.IO;

namespace SystemCounter
{
    public partial class Form1 : Form
    {                                                         
        int[] color_palette1 = {
            0x8ecae6,
            0x219ebc,
            0x023047,
            0xffb703,
            0xfb8500,
            0xf77f00,           // Orange
            0x003049,           // Dark blue
            0xd62828            // Red
        };

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

        public bool loadfile_isSuccess = false;
        double[] data           = { 1200, 1600, 1950, 2300, 2700 };
        string[] labels         = { "T_probe ANT1", "T_probe ANT2", "T_probe ANT3", "HDMI Connector", "RF Splitter" };
        int[] colors            = { 0x8ecae6, 0x219ebc, 0x023047, 0xf77f00, 0xd62828 };        
        int[] myData            = { 1200, 1600, 1950, 2300, 2700, 176, 3340, 2180, 1123, 546};

        public DataTable[] dataTable   = new DataTable[16];
        public DataTable idTable       = new DataTable();
        DataTable mTable        = new DataTable();
        public string[] chartTitles    = new string[16];

        DataTableCollection tableCollection;

        public static Form1 instance;

        TableReader tableReader = new TableReader();

        
        public Form1()
        {
            InitializeComponent();
            instance = this;
        }

        public string findStringBetween(string str, string startStr, string endStr)
        {
            int startIndex = str.IndexOf(startStr) + startStr.Length;
            int endIndex = str.IndexOf(endStr);
            string result = str.Substring(startIndex, endIndex - startIndex);
            return result;
        }

        
        public void update_chart()
        {
            idTable = DataForm.instance.idTable;
            dataTable = DataForm.instance.dataTable;

            tableReader.saveExcel(idTable, dataTable);

            //for (int i = 0; i < 16; i++)
            //    enable_chart(i, false);

            for(int i  = 0; i < idTable.Rows.Count; i++)
            {
                string deviceNum = idTable.Rows[i][0].ToString();
                string deviceId = idTable.Rows[i][1].ToString();

                string columnName = dataTable[i].Columns[0].ColumnName;
                Console.WriteLine(columnName);

                chartTitles[i] = deviceNum + " : " + deviceId;
                show_chart(i);
                printDataTable(dataTable[i]);
            }
            
            /*
            foreach (DataRow row in idTable.Rows)
            {
                string deviceNum = row["Device"].ToString();
                string deviceId = row["ID"].ToString();
                int num = Convert.ToInt16(deviceNum.Substring(deviceNum.Length - 2)) - 1;

                chartTitles[num] = deviceNum + " : " + deviceId;

                Console.WriteLine(chartTitles[num] +  " " + num);
                show_chart(num);
            }
            */
            
        }

        // ===================================================================== //
        //  Load Form
        // ===================================================================== //
        private void Form1_Load(object sender, EventArgs e)
        {
            
            // First-time open the form, need to check the settings.ini:
            string recordFilePath = get_ini_config("Settings", "FilePath");
            
            if (recordFilePath == "")
            {
                Console.WriteLine("The data is empty. Couldn't find the filepath!");
            }
            else
            {
                Console.WriteLine("> Load DataSet file: " + recordFilePath);
                load_recorded_data(recordFilePath);

                // Create an empty data table:
                for (int i = 0; i < 16; i++)
                {
                    dataTable[i] = new DataTable();
                    enable_chart(i, false);
                }

                
                foreach(DataRow row in idTable.Rows)
                {
                    string deviceNum = row["Device"].ToString();
                    string deviceId = row["ID"].ToString();
                    int num = Convert.ToInt16(deviceNum.Substring(deviceNum.Length - 2)) - 1;
                    
                    chartTitles[num] = deviceNum + " : " + deviceId;
                    dataTable[num] = get_device_data(deviceNum);

                    show_chart(num);
                }


                // Update Label:
                string stationName = get_ini_config("Settings", "StationName");
                label_title.Text = stationName;

                //load_excel(recordFilePath);
                //loadfile_isSuccess = true;
            }

            /*
            // If it successfully load the record.xlsx, show the chart:
            if (loadfile_isSuccess)
            {
                for(int i = 0; i < 16; i++)
                {
                    enable_chart(i, false);
                    if(dataTable[i].Rows.Count > 0)
                    {
                        enable_chart(i, true);
                        show_chart(i);
                    }
                }
                Console.WriteLine("> Successfully load the form!");
            }
            */

        







            /*
            // Load Data Table:
            //load_data();
            
            for (int i = 0; i < 16; i++)
            {
                dataTable[i] = new DataTable();
                enable_chart(i, false);
            }

            // Add id table empty value:
            idTable.Columns.Add("DeviceID");
            for(int i = 0; i < 16; i++)
            {
                idTable.Rows.Add("0000");
            }
            printDataTable(idTable);


            string path = get_ini_config("Settings", "FilePath");
            Console.WriteLine("FilePath="+path);
            
            /*
            load_excel("record.xlsx");
            for (int i = 0; i < 16; i++)
                show_chart(i);
            */
        }


        private DataTable get_device_data(string deviceNum)
        {
            DataTable table = new DataTable();
            try
            {
                table = tableCollection[deviceNum];
                return table;
            }
            catch { }
            return table;
        }


        // ===================================================================== //
        // Read Configuration File 
        // ===================================================================== //
        public void load_recorded_data(string path)
        {
            List<DataTable> dt = new List<DataTable>();
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
                        reader.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        // ====================================================================== //
        //  Load Data
        // ====================================================================== //
        private void load_data()
        { 
        }

        // ====================================================================== //
        // Load and Save the INI configuration File:
        // ====================================================================== //
        public string get_ini_config(string section, string key)
        {
            IniFile ini = new IniFile();
            string path = System.Environment.CurrentDirectory + "\\Settings.ini";
            string val = ini.ReadValue(path, section, key);
            return val;
        }
        public void set_ini_config(string section, string key, string val)
        {
            IniFile ini = new IniFile();
            string path = System.Environment.CurrentDirectory + "\\Settings.ini";
            ini.WriteValue(path, section, key, val);
        }


        // Read DataTable content:
        private string dataTable_getItem(DataTable table, string items, string column)
        {
            string val = "";            
            foreach(DataRow row in table.Rows)
            {
                string ret = row[0].ToString();
                if(ret == items)
                {
                    val = row[1].ToString();
                    Console.WriteLine("I found the string --> " + val);
                }
            }

            return val;
        }


        // Load excel dataset file:
        public void load_excel(string path)
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

                        for(int i = 0; i < 16; i++)
                        {
                            // Device ID:
                            string id = idTable.Rows[i][1].ToString();
                            string title = String.Format("Device {0} {1}", i + 1, id);
                            chartTitles[i] = title;

                            // Status of each device:
                            dataTable[i] = tableCollection[i + 1];                           
                            printDataTable(dataTable[i]);
                        }
                        reader.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

       




        private void getSamples(DataTable table, double[] data, string[] label)
        {
            // Prepare DataSet:
            int nSamples = table.Rows.Count;
            double[] chartData = new double[nSamples];
            string[] chartLabel = new string[nSamples];

            for (int i = 0; i < nSamples; i++)
            {
                string val = table.Rows[i]["Testing Time"].ToString();
                string testItem = table.Rows[i]["Items"].ToString();
                chartData[i] = Convert.ToDouble(val);
                chartLabel[i] = testItem;
            }

            data = chartData;
            label = chartLabel;

            double[] top5_data = new double[5];
            string[] top5_label = new string[5];

            for (int i = 0; i < 5; i++)
            {
                int startIndex = nSamples - 5;
                top5_data[i] = chartData[startIndex + i];
                top5_label[i] = chartLabel[startIndex + i];
            }
        }

        private void getTop5Samples(DataTable table, double[] data, string[] label)
        {
            // Prepare DataSet:
            int nSamples = table.Rows.Count;
            double[] chartData = new double[nSamples];
            string[] chartLabel = new string[nSamples];

            for (int i = 0; i < nSamples; i++)
            {
                string val = table.Rows[i]["Testing Time"].ToString();
                string testItem = table.Rows[i]["Items"].ToString();
                chartData[i] = Convert.ToDouble(val);
                chartLabel[i] = testItem;
            }
            
            double[] top5_data = new double[5];
            string[] top5_label = new string[5];

            for (int i = 0; i < 5; i++)
            {
                int startIndex = nSamples - 5;
                top5_data[i] = chartData[startIndex + i];
                top5_label[i] = chartLabel[startIndex + i];
            }
            data = top5_data;
            label = top5_label;

        }

        // ===================================================================== //
        //  Draw Chart
        // ===================================================================== //
        private void createChart(WinChartViewer viewer, string tittle, DataTable table)
        {           
            table.DefaultView.Sort = "Test Count";
            table = table.DefaultView.ToTable();


            // Prepare DataSet:
            int nSamples = table.Rows.Count;
            double[] chartData = new double[nSamples];
            string[] chartLabel = new string[nSamples];

            double[] top5_data = { 0, 0, 0, 0, 0};
            string[] top5_label = { "","","","",""};

            for (int i = 0; i < nSamples; i++)
            {
                string val = table.Rows[i]["Test Count"].ToString();
                string testItem = table.Rows[i]["Items"].ToString();
                chartData[i] = Convert.ToDouble(val);
                chartLabel[i] = testItem;
            }

            if (nSamples < 5)
            {
                int startIndex = 5 - nSamples;
                for(int i = 0; i < nSamples; i++)
                {
                    top5_data[startIndex + i] = chartData[i];
                    top5_label[startIndex + i] = chartLabel[i];
                }
            }
            else if (nSamples >= 5)
            {
                for (int i = 0; i < 5; i++)
                {
                    int startIndex = nSamples - 5;
                    top5_data[i] = chartData[startIndex + i];
                    top5_label[i] = chartLabel[startIndex + i];
                }
            }

            // Get chart viewer size:
            int height = viewer.Size.Height;
            int width = viewer.Size.Width;

            // Create chart
            XYChart c = new XYChart(width, height);
            c.setColor(ChartDirector.Chart.TextColor, 0xffffff);
            c.setBackground(0x000000);

            ChartDirector.TextBox title = c.addTitle(tittle, "Arial Bold", 14);
            title.setMargin2(0, 0, 6, 6);
            //c.addLine(20, title.getHeight(), c.getWidth() - 21, title.getHeight(), 0xffffff);

            c.setPlotArea(70, 80, 480, 240, -1, -1, ChartDirector.Chart.Transparent, ChartDirector.Chart.Transparent);
            c.swapXY();

            //c.addBarLayer3(data, colors).setBorderColor(Chart.Transparent);
            BarLayer layer = c.addBarLayer3(top5_data, colors);
            layer.setBorderColor(ChartDirector.Chart.Transparent);
            layer.setAggregateLabelFormat("{value}");
            
            c.xAxis().setLabels(labels);

            ChartDirector.TextBox textbox = c.xAxis().setLabels(top5_label);
            textbox.setFontStyle("Arial Bold Italic");
            textbox.setFontSize(10);
            textbox.setFontColor(0xffffff);
            //c.syncYAxis();

            //c.yAxis().setTitle("USD (millions)", "Arial Bold", 10);
            c.yAxis().setColors(ChartDirector.Chart.Transparent, ChartDirector.Chart.Transparent);
            c.yAxis2().setColors(ChartDirector.Chart.Transparent);
            c.xAxis().setTickColor(ChartDirector.Chart.Transparent);

            c.xAxis().setLabelStyle("Arial Bold", 8);
            //c.yAxis().setLabelStyle("Arial Bold", 8);
            //c.yAxis2().setLabelStyle("Arial Bold", 8);

            c.packPlotArea(30, title.getHeight() + 25, c.getWidth() - 50, c.getHeight() - 25);

            viewer.Chart = c;
            viewer.ImageMap = c.getHTMLImageMap("clickable", "", "title='{value}'");
        }


        // ===================================================================== //
        //  Print Data Table
        // ===================================================================== //
        public void printDataTable(DataTable table)
        {
            if (table.Rows.Count == 0) return; 

            try
            {
                Console.WriteLine("\n\n==================[ Print DataTable ]========================");
                //Console.WriteLine("Size : {0}x{1}\n", table.Rows.Count, table.Columns.Count);
                int row = table.Rows.Count;
                int col = table.Columns.Count;

                string str1 = table.Columns[0].ColumnName;
                string str2 = table.Columns[1].ColumnName;

                Console.WriteLine(str1 + "\t\t" + str2);

                foreach (DataRow data in table.Rows)
                {
                    for (int i = 0; i < col; i++)
                    {
                        Console.Write(data[i] + "\t\t");
                    }
                    Console.Write("\n");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("print_data_table: " + ex.Message);
            }
        }

       

        private void show_chart(WinChartViewer viewer)
        {
            Random rnd = new Random();

            double[] newData = new double[5];
            for(int i = 0; i < 5; i++)
                newData[i] = rnd.Next(100, 3000);
             
            int height  = viewer.Size.Height;
            int width   = viewer.Size.Width;

            XYChart c = new XYChart(width, height);
            c.setColor(ChartDirector.Chart.TextColor, 0xffffff);
            c.setBackground(0x000000);
           
            ChartDirector.TextBox title = c.addTitle("Device 1 3370302", "Arial Bold", 14);
            title.setMargin2(0, 0, 6, 6);
            //c.addLine(20, title.getHeight(), c.getWidth() - 21, title.getHeight(), 0xffffff);

            c.setPlotArea(70, 80, 480, 240, -1, -1, ChartDirector.Chart.Transparent, ChartDirector.Chart.Transparent);
            c.swapXY();

            //c.addBarLayer3(data, colors).setBorderColor(Chart.Transparent);
            BarLayer layer = c.addBarLayer3(newData, colors);
            layer.setBorderColor(ChartDirector.Chart.Transparent);
            layer.setAggregateLabelFormat("{value}");
            

            c.xAxis().setLabels(labels);

            ChartDirector.TextBox textbox = c.xAxis().setLabels(labels);
            textbox.setFontStyle("Arial Bold Italic");
            textbox.setFontSize(10);
            textbox.setFontColor(0xffffff);
            //c.syncYAxis();
            //c.yAxis().setTitle("USD (millions)", "Arial Bold", 10);

            c.yAxis().setColors(ChartDirector.Chart.Transparent, ChartDirector.Chart.Transparent);
            c.yAxis2().setColors(ChartDirector.Chart.Transparent);
            c.xAxis().setTickColor(ChartDirector.Chart.Transparent);

            c.xAxis().setLabelStyle("Arial Bold", 8);
            //c.yAxis().setLabelStyle("Arial Bold", 8);
            //c.yAxis2().setLabelStyle("Arial Bold", 8);

            c.packPlotArea(30, title.getHeight() + 25, c.getWidth() - 50, c.getHeight() - 25);

            viewer.Chart = c;
            viewer.ImageMap = c.getHTMLImageMap("clickable", "", "title='{value}'");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            /*
            Form2 newForm = new Form2();
            newForm.ShowDialog();
            */
        }

        private void button2_Click(object sender, EventArgs e)
        {          
            DataForm form = new DataForm();
            form.dataTable = dataTable;
            form.idTable = idTable;
            form.ShowDialog();            
        }

        private void winChartViewer1_SizeChanged(object sender, EventArgs e)
        {
            
        }

        private void winChartViewer1_SizeChanged_1(object sender, EventArgs e)
        {
            
        }
       
        public void enable_chart(int id, bool state)
        {
            
            switch (id)
            {
                case 0:
                    winChartViewer1.Visible = state;
                    break;
                case 1:
                    winChartViewer2.Visible = state;
                    break;
                case 2:
                    winChartViewer3.Visible = state;
                    break;
                case 3:
                    winChartViewer4.Visible = state;
                    break;
                case 4:
                    winChartViewer5.Visible = state;
                    break;
                case 5:
                    winChartViewer6.Visible = state;
                    break;
                case 6:
                    winChartViewer7.Visible = state;
                    break;
                case 7:
                    winChartViewer8.Visible = state;
                    break;
                case 8:
                    winChartViewer9.Visible = state;
                    break;
                case 9:
                    winChartViewer10.Visible = state;
                    break;
                case 10:
                    winChartViewer11.Visible = state;
                    break;
                case 11:
                    winChartViewer12.Visible = state;
                    break;
                case 12:
                    winChartViewer13.Visible = state;
                    break;
                case 13:
                    winChartViewer14.Visible = state;
                    break;
                case 14:
                    winChartViewer15.Visible = state;
                    break;
                case 15:
                    winChartViewer16.Visible = state;
                    break;
            }
        }

        private void show_chart(int deviceid)
        {
           
            enable_chart(deviceid, true);
            DataTable table = dataTable[deviceid];

            try {
                switch (deviceid)
                {
                    case 0:
                        createChart(winChartViewer1, chartTitles[deviceid], table);
                        break;
                    case 1:
                        createChart(winChartViewer2, chartTitles[deviceid], table);
                        break;
                    case 2:
                        createChart(winChartViewer3, chartTitles[deviceid], table);
                        break;
                    case 3:
                        createChart(winChartViewer4, chartTitles[deviceid], table);
                        break;
                    case 4:
                        createChart(winChartViewer5, chartTitles[deviceid], table);
                        break;
                    case 5:
                        createChart(winChartViewer6, chartTitles[deviceid], table);
                        break;
                    case 6:
                        createChart(winChartViewer7, chartTitles[deviceid], table);
                        break;
                    case 7:
                        createChart(winChartViewer8, chartTitles[deviceid], table);
                        break;
                    case 8:
                        createChart(winChartViewer9, chartTitles[deviceid], table);
                        break;
                    case 9:
                        createChart(winChartViewer10, chartTitles[deviceid], table);
                        break;
                    case 10:
                        createChart(winChartViewer11, chartTitles[deviceid], table);
                        break;
                    case 11:
                        createChart(winChartViewer12, chartTitles[deviceid], table);
                        break;
                    case 12:
                        createChart(winChartViewer13, chartTitles[deviceid], table);
                        break;
                    case 13:
                        createChart(winChartViewer14, chartTitles[deviceid], table);
                        break;
                    case 14:
                        createChart(winChartViewer15, chartTitles[deviceid], table);
                        break;
                    case 15:
                        createChart(winChartViewer16, chartTitles[deviceid], table);
                        break;
                }
            }
            catch(Exception ex) {
                Console.WriteLine(ex.Message);
            }
                
        }


        // Show the table in chart
        private void button1_Click(object sender, EventArgs e)
        {
            /*
            DataTable table = new DataTable();
            table.Columns.Add("Items", typeof(string));
            table.Columns.Add("Testing Time", typeof(int));
            table.Rows.Add("T-Probe1", 178);
            table.Rows.Add("T-Probe2", 2889);
            table.Rows.Add("T-Probe3", 411);
            table.Rows.Add("T-Probe4", 4028);
            table.Rows.Add("Splitter1", 1200);
            table.Rows.Add("Splitter2", 1480);
            table.Rows.Add("Splitter3", 3267);
            table.Rows.Add("SMA-Cable1", 867);
            table.Rows.Add("SMA-Cable2", 868);
            table.Rows.Add("SMA-Cable3", 980);

            // Sort the table based on the value
            table.DefaultView.Sort = "Testing Time";
            table = table.DefaultView.ToTable();

            mTable = table;
            
            createChart(winChartViewer1, chartTitles[0], mTable);
            printDataTable(table);        
            */
        }




        private void winChartViewer1_Click_1(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[0], dataTable[0]);
            myForm.ShowDialog();
        }
        private void winChartViewer2_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[1], dataTable[1]);
            myForm.ShowDialog();
        }

        private void winChartViewer3_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[2], dataTable[2]);
            myForm.ShowDialog();
        }

        private void winChartViewer4_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[3], dataTable[3]);
            myForm.ShowDialog();
        }

        private void winChartViewer5_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[4], dataTable[4]);
            myForm.ShowDialog();
        }

        private void winChartViewer6_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[5], dataTable[5]);
            myForm.ShowDialog();
        }

        private void winChartViewer7_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[6], dataTable[6]);
            myForm.ShowDialog();
        }

        private void winChartViewer8_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[7], dataTable[7]);
            myForm.ShowDialog();
        }

        private void winChartViewer9_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[8], dataTable[8]);
            myForm.ShowDialog();
        }

        private void winChartViewer10_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[9], dataTable[9]);
            myForm.ShowDialog();
        }

        private void winChartViewer11_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[10], dataTable[10]);
            myForm.ShowDialog();
        }

        private void winChartViewer12_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[11], dataTable[11]);
            myForm.ShowDialog();
        }

        private void winChartViewer13_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[12], dataTable[12]);
            myForm.ShowDialog();
        }

        private void winChartViewer14_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[13], dataTable[13]);
            myForm.ShowDialog();
        }

        private void winChartViewer15_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[14], dataTable[14]);
            myForm.ShowDialog();
        }

        private void winChartViewer16_Click(object sender, EventArgs e)
        {
            LargeForm myForm = new LargeForm();
            myForm.StartPosition = FormStartPosition.CenterParent;
            myForm.setChartData(chartTitles[15], dataTable[15]);
            myForm.ShowDialog();
        }

        // ====================================================================== //
        private void winChartViewer1_SizeChanged_3(object sender, EventArgs e)
        {
            if (dataTable[0].Rows.Count != 0)
                createChart(winChartViewer1, chartTitles[0], dataTable[0]);               
        }

        private void winChartViewer2_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[1].Rows.Count != 0)
                createChart(winChartViewer2, chartTitles[1], dataTable[1]);
        }

        private void winChartViewer3_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[2].Rows.Count != 0)
                createChart(winChartViewer3, chartTitles[2], dataTable[2]);
        }

        private void winChartViewer4_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[3].Rows.Count != 0)
                createChart(winChartViewer4, chartTitles[3], dataTable[3]);
        }

        private void winChartViewer5_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[4].Rows.Count != 0)
                createChart(winChartViewer5, chartTitles[4], dataTable[4]);
        }

        private void winChartViewer6_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[5].Rows.Count != 0)
                createChart(winChartViewer6, chartTitles[5], dataTable[5]);
        }

        private void winChartViewer7_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[6].Rows.Count != 0)
                createChart(winChartViewer7, chartTitles[6], dataTable[6]);
        }

        private void winChartViewer8_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[7].Rows.Count != 0)
                createChart(winChartViewer8, chartTitles[7], dataTable[7]);
        }

        private void winChartViewer9_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[8].Rows.Count != 0)
                createChart(winChartViewer9, chartTitles[8], dataTable[8]);
        }

        private void winChartViewer10_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[9].Rows.Count != 0)
                createChart(winChartViewer10, chartTitles[9], dataTable[9]);
        }

        private void winChartViewer11_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[10].Rows.Count != 0)
                createChart(winChartViewer11, chartTitles[10], dataTable[10]);
        }

        private void winChartViewer12_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[11].Rows.Count != 0)
                createChart(winChartViewer12, chartTitles[11], dataTable[11]);
        }

        private void winChartViewer13_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[12].Rows.Count != 0)
                createChart(winChartViewer13, chartTitles[12], dataTable[12]);
        }

        private void winChartViewer14_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[13].Rows.Count != 0)
                createChart(winChartViewer14, chartTitles[13], dataTable[13]);
        }

        private void winChartViewer15_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[14].Rows.Count != 0)
                createChart(winChartViewer15, chartTitles[14], dataTable[14]);
        }

        private void winChartViewer16_SizeChanged(object sender, EventArgs e)
        {
            if (dataTable[15].Rows.Count != 0)
                createChart(winChartViewer16, chartTitles[15], dataTable[15]);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            System.Data.DataSet tables = new System.Data.DataSet();
            bool ret = tableReader.readExcel("record.xlsx", ref tables);
            if (!ret)
            {
                Console.WriteLine("Failed to open the file!");
                return;
            }
            DataTable myTable = tables.Tables["DeviceID"];
            tableReader.print(myTable);

            System.Data.DataSet dataSet = new System.Data.DataSet();
            
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            TitleForm form = new TitleForm();
            form.ShowDialog();
        }
    }
}
