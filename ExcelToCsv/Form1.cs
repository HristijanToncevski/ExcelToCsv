using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace ExcelToCsv
{
    public partial class ExcelToCsv : Form
    {
        public ExcelToCsv()
        {
            InitializeComponent();
        }
     
        private void btnChoose_Click(object sender, EventArgs e)
        {
            ExcelApp.Application excelApp = new ExcelApp.Application();
            ExcelApp.Workbook excelBook = excelApp.Workbooks.Open("C:\\CSV\\Import.xlsx"); // Excel Document
            ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
            ExcelApp.Range excelRange = excelSheet.UsedRange;

            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            dataGridView1.RowCount = rows;
            dataGridView1.ColumnCount = cols;

            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        dataGridView1.Rows[i - 1].Cells[j - 1].Value = excelRange.Cells[i, j].Value2.ToString();
                }
            }
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            string csv = string.Empty;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    csv += "\"" + cell.Value + "\"";
                    csv += ",";
                }
                csv += "\"\",\"\",\"\"";
                csv += "\r\n";
            }
            string result = csv.Remove(csv.TrimEnd().LastIndexOf(Environment.NewLine));
            string folderPath = "C:\\CSV\\"; // CSV Export Document
            File.WriteAllText(folderPath + "Export.csv", result);
        }
    }
}
