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
        private List<String> departments = new List<string>();
        private List<String> jobTitles = new List<string>();

        public void populateTOC()
        {
            //Add Table Of Contentds page and push to front of workbook
            excelFile.Workbook.Worksheets.Add("Table Of Contents");
            excelFile.Workbook.Worksheets.MoveToStart("Table Of Contents");
            ExcelWorksheet tableofContents = excelFile.Workbook.Worksheets["Table Of Contents"];

            //Set column headers
            tableofContents.Cells[1, 1].Value = "Job Titles";
            tableofContents.Cells[1, 3].Value = "Departments";

            tableofContents.Cells[1, 1, 1, 10].Style.Font.Bold = true;
            tableofContents.Cells[1, 1, 1, 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            //Create a new style for our hyperlinks
            var namedStyle = tableofContents.Workbook.Styles.CreateNamedStyle("HyperLink"); 
            namedStyle.Style.Font.UnderLine = true;
            namedStyle.Style.Font.Color.SetColor(Color.Blue);

            //Populate respective columns with hyperlinks
            int tracker = 2;
            List<string> jobs = getDepartments("Job Title");
            foreach (string title in jobs)
            { 
                tableofContents.Cells[tracker, 1].Hyperlink = new ExcelHyperLink("'" + title + "'" + "!A1", title);
                tableofContents.Cells[tracker, 1].StyleName = "HyperLink";
                tracker++;
            }

            tracker = 2;
            List<string> Departments = getDepartments("Departments");
            foreach (string dept in Departments)
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
