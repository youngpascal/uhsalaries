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
        Data ld;

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

        }//end form load

        public String getFilePath()
        {
            return filePath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int constantL = int.Parse(textBox1.Text);
            int constantD = int.Parse(textBox2.Text);
            int constantK = int.Parse(textBox3.Text);

            ld = new Data(filePath);
           // ld.load(2);
            //ld.load(12);
            //ld.load(27);

            //ld.load(2);
            //ld.searchForHeader("Job Title");
            //ld.searchForHeader("Pos Deptid");
            //ld.searchForHeader("Total Salary");

            ld.prepareData("FY2013 Detail Faculty Roster", "Prepared Data");
            Console.WriteLine("Done copying");
            ld.sortPreparedData();
            Console.WriteLine("Done sorting and populating");
            //ld.sortPreparedDeptData();
            //Console.WriteLine("Done sorting departments");
        }


    }
}