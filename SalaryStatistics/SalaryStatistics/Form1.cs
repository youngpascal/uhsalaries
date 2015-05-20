using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web.UI.WebControls;
using System.IO;
//EPPlus API
using OfficeOpenXml;

namespace SalaryStatistics
{
    public partial class Form1 : Form
    {
        private String fiscalFilePath = ""; // @"C:\Users\josh\Desktop\Salary inputs\Input 0.xlsx";
        private String inputOneFilePath =  ""; // @"C:\Users\josh\Desktop\Salary inputs\Input 1.xlsx";
        private String inputTwoFilePath = ""; //@"C:\Users\josh\Desktop\Salary inputs\Input 2.xlsx";
        private String inputThreeFilePath = "";//@"C:\Users\josh\Desktop\Salary inputs\Input 3.xlsx";
        private Data theData;        
        Form1 myForm;

        public Form1()
        {
            InitializeComponent();
           // this.TopMost = true;
        }

        //Prompts a user to select the input file before displaying the settings window with the action button.
        private void Form1_Load(object sender, EventArgs e)
        {
            //If the log file doesn't exist, create it. requires run as admin
            //if (!File.Exists(@"C:\salaryLog.txt"))
            {
                //FileStream fs = new FileStream(@"C:\salaryLog.txt", FileMode.CreateNew);
               // fs.Close();
            }
            button6.Visible = false;
            label5.Text = fiscalFilePath;
            label6.Text = inputOneFilePath;
            label7.Text = inputTwoFilePath;
            label8.Text = inputThreeFilePath;
        }//end form load

        //Getter for From1's filePath string.
        public String getFilePath()
        {
            return fiscalFilePath;
        }

        //Triggered by the action button. Takes the constants specified and the filepath to create a Data object.
        //Then it applys the prepare(), process(), and close() fuctions to the object.
        private void button1_Click(object sender, EventArgs e)
        {
            doWork();
        }

        private void getFilters()
        {
            float constantL = float.Parse(textBox1.Text);
            float constantD = float.Parse(textBox3.Text);
            float constantK = float.Parse(textBox2.Text);
            float constantL2 = float.Parse(textBox4.Text);
            string sourceSheetName = "0"; // "FY2013 Detail Faculty Roster";
            string preparedSheetName = "Prepared Data";
            

            theData = new Data(fiscalFilePath, inputOneFilePath, inputTwoFilePath, inputThreeFilePath, constantD, constantK, constantL, constantL2);

            theData.Prepare(sourceSheetName, preparedSheetName, "Job Title");
            theData.copyInputOne();
            theData.copyInputTwo();
            theData.copyInputThree();

            ExcelWorksheet prepared = theData.getExcelFile().Workbook.Worksheets[preparedSheetName];

            //Get all the possible filters in the worksheet
            Dictionary<string, int> keyColumns = theData.getKeyColumns();
            theData.fetchFilters(keyColumns, prepared);

            List<string> filters = theData.getJobFiltersList();
            List<string> departmentFilters = theData.getDepartmentFiltersList();

            //Sort the Lists
            filters.Sort();
            departmentFilters.Sort();

            //Populate the check box lists
            checkedFilters.Items.AddRange(filters.ToArray());
            checkedDepartmentFilters.Items.AddRange(departmentFilters.ToArray());
        }

        private void doWork()
        {
            float constantL = float.Parse(textBox1.Text);
            float constantD = float.Parse(textBox3.Text);
            float constantK = float.Parse(textBox2.Text);
            float constantL2 = float.Parse(textBox4.Text);
            bool filtered = false;
            int jobFilterCount = 0;
            int departmentFilterCount = 0;
            List<string> searchFilters = new List<string>();

            getFilters();
            //Get all of the checked filters
            foreach (Object list in checkedFilters.CheckedItems)
            {
                searchFilters.Add(list.ToString());
                jobFilterCount++;
            }

            foreach (Object list in checkedDepartmentFilters.CheckedItems)
            {
                searchFilters.Add(list.ToString());
                departmentFilterCount++;
            }
            
            //Determine if filters were selected
            if (searchFilters.Any())
            {
               filtered = true;
            }

            theData.setJobFilterCount(jobFilterCount);
            theData.setDepartmentFilterCount(departmentFilterCount);

            //Process and close
            Cursor.Current = Cursors.WaitCursor;

            theData.Process(searchFilters, filtered);
            theData.populateTOC();
            theData.specialtyCodes();
            theData.Close();

            Cursor.Current = Cursors.Default;
            //End process and close

        }

        private string selectExcelSheets(string title)
        {
            Stream myStream = null;

            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = @"C:\Users\%USERNAME%";
            of.Filter = "Excel files (*.xls*)|*.xls*|All files (*.*)|*.*";
            of.FilterIndex = 1;
            of.Title = title;
            of.RestoreDirectory = true;
            string filePath = "";

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
            }//end if dialogresult=ok

            return filePath;
        }

        //import input 0
        private void button2_Click(object sender, EventArgs e)
        {
            fiscalFilePath = selectExcelSheets("Select excel document containing UH fiscal year data.");
            label5.Text = fiscalFilePath;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            inputOneFilePath = selectExcelSheets("Select 'input 1' excel document containing Average New Assistant Professor Salaries.");
            label6.Text = inputOneFilePath;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            inputTwoFilePath = selectExcelSheets("Select 'input 2' excel document containing Tier 1 data.");
            label7.Text = inputTwoFilePath;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            inputThreeFilePath = selectExcelSheets("Select 'inout 3' excel document containing UH data per specialty code.");
            label8.Text = inputThreeFilePath;
            //button6.Visible = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            getFilters();
        }
    }
}