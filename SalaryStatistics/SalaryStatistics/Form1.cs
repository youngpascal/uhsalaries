using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
//Ep Plus API
using OfficeOpenXml;

namespace SalaryStatistics
{
    public partial class Form1 : Form
    {
        private String filePath = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Stream myStream = null;
            
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = "c:\\";
            //of.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            //of.FilterIndex = 2;
            of.RestoreDirectory = true;

            if (of.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = of.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            //Store file path in filePath
                            filePath = of.FileName;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }

               // myStream.Close();
            }//end if dialogresult=ok
                FileInfo existingFile = new FileInfo(filePath);
                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    // get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    int deletedRows = 0;
                    int col = 2;

                    //Delete all the extra rows in the first 
                    for (int row = 1; row < 51; row++)
                    {
                        if (worksheet.Cells[row, col].Value == null)
                        {
                            worksheet.DeleteRow(row, 1, true);
                            deletedRows++;
                        }
                    }

                    for (int row = 1; row < 51 - deletedRows; row++ )
                        Console.WriteLine("\tCell({0},{1}).Value={2}", row, col, worksheet.Cells[row, col].Value);

                    Console.WriteLine("\tDeleted Rows: {0}", deletedRows);
                } // the using statement automatically calls Dispose() which closes the package.
        }//end form load

        public String getFilePath()
        {
            return filePath;
        }


    }
}
