using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Windows.Forms;
using ExcelDataReader;
using Bytescout.Spreadsheet;

namespace SystemCounter
{
    class TableReader
    {
        public void print(DataTable table)
        {
            //if (table.Rows.Count == 0) return;
            try
            {
                Console.WriteLine("\n\n==================[ Print DataTable ]========================");
                //Console.WriteLine("Size : {0}x{1}\n", table.Rows.Count, table.Columns.Count);
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
            }
            catch (Exception ex)
            {
                MessageBox.Show("print_data_table: " + ex.Message);
            }
        }

        public bool readExcel(string path, ref DataSet tables)
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
                        tables = result;
                        reader.Close();
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return false;
        }

        
        private void tableToSheet(Spreadsheet document, ref Spreadsheet newDocument,string sheetName, DataTable table)
        {
            try {
                Worksheet sheet = document.Workbook.Worksheets.Add(sheetName);
                int nRow = table.Rows.Count;
                int nCol = table.Columns.Count;

                sheet.Cell(0, 0).Value = table.Columns[0].ColumnName;
                sheet.Cell(0, 1).Value = table.Columns[1].ColumnName;

                for (int i = 0; i < nRow; i++)
                {
                    for (int j = 0; j < nCol; j++)
                    {
                        sheet.Cell(i + 1, j).Value = table.Rows[i][j];
                    }
                }
                newDocument = document;
            }
            catch(Exception ex) {
                Console.WriteLine(ex.Message);
            }
        }

        public bool saveExcel(DataTable idTables, DataTable[] dataTable)
        {
            Spreadsheet document = new Spreadsheet();
            tableToSheet(document, ref document, "DeviceID", idTables);

            for(int i = 0; i < 16; i++)
            {
                if (dataTable[i].Rows.Count != 0)
                {
                    string sheetName = "Device " + (i + 1);
                    tableToSheet(document, ref document, sheetName, dataTable[i]);
                }
            }

            document.SaveAs("record.xlsx");

            return false;
        }

        public void print_allTable(DataTable[] table)
        {
            Console.WriteLine("Total Table = " + table.Length);
            print(table[0]);
            print(table[1]);
            print(table[2]);
            print(table[3]);
            print(table[4]);
            print(table[5]);
        }


        // Set table values:
        public void seValue(DataTable inTable, ref DataTable outTable, string item, int val)
        {
            outTable = inTable;
            int nRow = inTable.Rows.Count;
            for(int i = 0; i < nRow; i++)
            {
                if(inTable.Rows[i][0].ToString() == item)
                {
                    outTable.Rows[i][1] = val;
                }
            }
        }
        
        public int getValue(DataTable inTable, string item)
        {
            int val = 0;
            int nRow = inTable.Rows.Count;
            for (int i = 0; i < nRow; i++)
            {
                if (inTable.Rows[i][0].ToString() == item)
                {
                    string val_str = inTable.Rows[i][1].ToString();
                    if (val_str != "")
                        val = Convert.ToInt16(val_str);
                }
            }
            return val;
        }
    }
}
