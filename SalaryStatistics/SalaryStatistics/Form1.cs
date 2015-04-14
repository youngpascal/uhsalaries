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
                loadData ld = new loadData();
                DirectoryInfo di = new DirectoryInfo(filePath);
                DataTable dt = ld.GetDataTable(di);

                DataRow[] foundRows;
                foundRows = dt.Select("Job Title");
                MessageBox.Show(" ");
                for (int i = 0; i < foundRows.Length; i++)
                {
                    Console.WriteLine(foundRows[i][0]);
                }
        }//end form load

        public String getFilePath()
        {
            return filePath;
        }


    }
}
