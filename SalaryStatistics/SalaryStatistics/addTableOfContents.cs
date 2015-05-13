using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Drawing;

namespace SalaryStatistics
{
    public partial class Data
    {

        public void populateTOC()
        {
            //Temporary Lists for creating hyperlinks
            List<String> departments = new List<string>();
            List<String> jobTitles = new List<string>();

            //Add Table Of Contentds page and push to front of workbook
            excelFile.Workbook.Worksheets.Add("Table Of Contents");
            excelFile.Workbook.Worksheets.MoveToStart("Table Of Contents");
            ExcelWorksheet tableofContents = excelFile.Workbook.Worksheets["Table Of Contents"];

            //Set column headers
            tableofContents.Cells[1, 1].Value = "Job Titles";
            tableofContents.Cells[1, 3].Value = "Departments";

            //Create a new style for our hyperlinks
            var namedStyle = tableofContents.Workbook.Styles.CreateNamedStyle("HyperLink"); 
            namedStyle.Style.Font.UnderLine = true;
            namedStyle.Style.Font.Color.SetColor(Color.Blue);

            //Seperate all worksheet names by department and job title
            foreach (string name in processedWorksheetNames)
            {
                if (name[0].Equals('H') || name[0].Equals('h'))
                {
                    departments.Add(name);
                }
                else
                {
                    jobTitles.Add(name);
                }
            }

            //Sort the lists by alphabetical order
            jobTitles.Sort();
            departments.Sort();

            //Populate respective columns with hyperlinks
            int tracker = 2;

            foreach (string title in jobTitles)
            {
                
                tableofContents.Cells[tracker, 1].Hyperlink = new ExcelHyperLink("'" + title + "'" + "!A1", title);
                tableofContents.Cells[tracker, 1].StyleName = "HyperLink";
                tracker++;
            }

            tracker = 2;

            foreach (string dept in departments)
            {
                tableofContents.Cells[tracker, 3].Hyperlink = new ExcelHyperLink("'" + dept + "'" + "!A1", dept);
                tableofContents.Cells[tracker, 3].StyleName = "HyperLink";
                tracker++;
            }

            //Adust column width
            tableofContents.Cells["A:Z"].AutoFitColumns();
        }    
    }
}
