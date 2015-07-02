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
        public void specialtyCodes()
        {
            ExcelWorksheet uhSpecialties = excelFile.Workbook.Worksheets["UH Specialty Averages"];
            ExcelWorksheet tierOneSpecialties = excelFile.Workbook.Worksheets["Tier 1 Data"];

            List<uhSpecialties> uhSpecialtyList = new List<uhSpecialties>();
            List<tierOneSpecialties> tierOneSpecialtyList = new List<tierOneSpecialties>();

            string spec;
            string weight;
            string avg = "";
            string jt = "";
            string dept = "";

            for (int row = 2; row <= uhSpecialties.Dimension.End.Row; row++)
            {
                spec = uhSpecialties.Cells[row, 1].Value.ToString();
                weight = uhSpecialties.Cells[row, 4].FullAddress;
                jt = uhSpecialties.Cells[row, 2].Value.ToString();
                dept = uhSpecialties.Cells[row, 3].Value.ToString();
                uhSpecialtyList.Add(new uhSpecialties(spec, weight, jt, dept));
            }

            for (int row = 2; row <= tierOneSpecialties.Dimension.End.Row; row++)
            {
                avg = tierOneSpecialties.Cells[row, 1].FullAddress;
                jt = tierOneSpecialties.Cells[row, 2].Value.ToString();
                spec = tierOneSpecialties.Cells[row, 3].Value.ToString();
                tierOneSpecialtyList.Add(new tierOneSpecialties(spec, avg, jt));
            }

            //Create the list of Tier 1 Average salaries, adjusted for the specalty makeups of UH's departments
            ExcelWorksheet weightedDeptPortions = excelFile.Workbook.Worksheets.Add("Weighted Dept Salary Portions");
            weightedDeptPortions.Cells[1, 1].Value = "Job Title";
            weightedDeptPortions.Cells[1, 2].Value = "Department";
            weightedDeptPortions.Cells[1, 3].Value = "Specialty Code";
            weightedDeptPortions.Cells[1, 4].Value = "Weighted Salary Portion";

            //Match structs by specialty code and job title, then pull the average salary address
            int currentRow = 2;
            foreach (uhSpecialties uh in uhSpecialtyList)
            {
                weightedDeptPortions.Cells[currentRow, 1].Value = uh.jobTitle;
                weightedDeptPortions.Cells[currentRow, 2].Value = uh.department;
                weightedDeptPortions.Cells[currentRow, 3].Value = uh.specialtyCode;

                foreach(tierOneSpecialties tierOne in tierOneSpecialtyList)
                {
                    if (uh.specialtyCode.Equals(tierOne.specialtyCode) && uh.jobTitle.Equals(tierOne.jobTitle))
                    {
                        weightedDeptPortions.Cells[currentRow, 4].Formula = uh.weight + "*" + tierOne.averageSalary;
                        break;
                    }
                }

                currentRow++;
            }

            weightedDeptPortions.Cells["D:D"].Style.Numberformat.Format = "$###,###,##0";
            weightedDeptPortions.Cells["A:Z"].AutoFitColumns();


            //Add the sheet listing the university-wide average for each job title
            ExcelWorksheet weightedPositionAvg = excelFile.Workbook.Worksheets.Add("Job Title Weighted Salary Avg.");
            weightedPositionAvg.Cells[1, 1].Value = "Job Title";
            weightedPositionAvg.Cells[1, 2].Value = "Department";
            weightedPositionAvg.Cells[1, 3].Value = "UH Average";
            weightedPositionAvg.Cells[1, 4].Value = "Tier 1 Weighted Average";
            weightedPositionAvg.Cells[1, 5].Value = "Ratio";

            currentRow = 2;
            int fromRow;

            foreach (ExcelWorksheet worksheet in excelFile.Workbook.Worksheets)
            {
                if (worksheet.Name[0].ToString().ToUpper() == "H" && worksheet.Dimension.End.Row > 3)
                {
                    fromRow = 3;
                    //Loop through all summary rows and copy data over, exit when there's a data row
                    while (worksheet.Cells[fromRow, 1].Value != null)
                    {
                        weightedPositionAvg.Cells[currentRow, 1].Value = worksheet.Cells[fromRow, 1].Value; //Job Title
                        weightedPositionAvg.Cells[currentRow, 2].Value = worksheet.Name; //Department
                        weightedPositionAvg.Cells[currentRow, 3].Formula = "+" + worksheet.Name + "!E" + fromRow; //UH Average

                        fromRow++;
                        currentRow++;
                    }
                }
            }

            currentRow = 2;
            string formula;
            //Loop through all the Tier 1 specalty code portions and set the appropriate summing formulas
            while (currentRow <= weightedPositionAvg.Dimension.End.Row)
            {
                formula = "=0";
                fromRow = 2;
                while (fromRow <= weightedDeptPortions.Dimension.End.Row)
                {
                    if (weightedDeptPortions.Cells[fromRow, 1].Value.ToString().ToUpper() == weightedPositionAvg.Cells[currentRow, 1].Value.ToString().ToUpper()
                        && weightedDeptPortions.Cells[fromRow, 2].Value.ToString().ToUpper() == weightedPositionAvg.Cells[currentRow, 2].Value.ToString().ToUpper())
                    {
                        formula += "+" + "'" + weightedDeptPortions.Name + "'" + "!D" + fromRow;
                        //Doing the formula as adding specific cells and not a sum of adjecent cells makes this impervious to sorting of weightedDeptProtions' rows.
                    }
                    fromRow++;
                }
                if (formula != "=0")
                {
                    weightedPositionAvg.Cells[currentRow, 4].Formula = formula; //Set the formula if one was generated
                    weightedPositionAvg.Cells[currentRow, 5].Formula = "=C" + currentRow + "/D" + currentRow; //Set Ratio
                }
                currentRow++;
            }

            /* //Old code, partially modified (not in a working state)
            currentRow = 2;
            List<string> jobTitles = getLists("Job Title");
            foreach (string jobTitle in jobTitles)
            {   //Set the average for that jobTitle
                weightedPositionAvg.Cells[currentRow, 1].Value = jobTitle;
                weightedPositionAvg.Cells[currentRow, 2].Formula = "'" + jobTitle + "'" + "!E2";

                int firstRow = 0;
                int counter = 0;
                bool toBreak = false;
                bool wasFound = false;
                //Loop through all rows in the sheet of job title weights as a portion of the total jobs in the department.
                for (int row = 2; row <= weightedDeptPortions.Dimension.End.Row; row++)
                {
                    firstRow = row;
                    //Loop through 
                    while (weightedDeptPortions.Cells["B" + row].Value.ToString().ToUpper().Equals(jobTitle.ToString().ToUpper()))
                    {
                        counter++;
                        if (!weightedDeptPortions.Cells["B" + (row + 1)].Value.ToString().Equals(jobTitle))
                        {
                            toBreak = true;
                            wasFound = true;
                            break;
                        }
                        row++;
                    }
                    if (toBreak)
                    {
                        break;
                    }
                }
                if (wasFound)
                {
                    weightedPositionAvg.Cells[currentRow, 3].Formula = "SUM('Weighted Dept Salary Portions'!D" + firstRow + ":" + "'Weighted Dept Salary Portions'!D" + (firstRow + counter - 1) + ")";
                }
                else
                {
                    weightedPositionAvg.Cells[currentRow, 3].Value = "Department not found";
                }
                weightedPositionAvg.Cells[currentRow, 4].Formula = "C" + currentRow + "/" + "B" + currentRow;
                currentRow++;
            }
            */

            weightedPositionAvg.Cells["B:D"].Style.Numberformat.Format = "$###,###,##0";
            //weightedPositionAvg.Cells["E"].Style.Numberformat.Format = "0.00";
            weightedPositionAvg.Cells["A:Z"].AutoFitColumns();


            ExcelWorksheet statisticsSummary = excelFile.Workbook.Worksheets.Add("Department per Job Title Statistics Summary");

            string[] statisticHeaders = {"Job Title", "Department" , "Zeros", "1st Quartile", "Mean", "3rd Quartile", "Median", "Associate Prof. Compression" };
            //Insert the statistic column headers
            for (int r = 0; r < 8; r++)
            {
                statisticsSummary.Cells[1, r + 1].Value = statisticHeaders[r];
            }

            currentRow = 3;
            foreach (ExcelWorksheet worksheet in excelFile.Workbook.Worksheets)
            {
                if (worksheet.Name[0].ToString().ToUpper() == "H" && worksheet.Dimension.End.Row > 3)
                {
                    fromRow = 3;
                    //Loop through all summary rows and copy data over, exit when there's a data row
                    while (worksheet.Cells[fromRow, 1].Value != null)
                    {
                        for (int i = 1; i < 9; i++) {
                            statisticsSummary.Cells[currentRow, i].Formula = "=" + worksheet.Cells[fromRow, i].FullAddress; //Job Title
                        }

                        statisticsSummary.Cells[currentRow, 2].Value = worksheet.Name; //Set Department, since it's not in the summary rows.

                        fromRow++;
                        currentRow++;
                    } 
                }
            }

            statisticsSummary.Cells["E:G"].Style.Numberformat.Format = "$###,###,##0";
            statisticsSummary.Cells["A:Z"].AutoFitColumns();


            //Create the Table of Contents
            ExcelWorksheet toc = excelFile.Workbook.Worksheets["Table Of Contents"];

            toc.Cells[1, 5].Value = "Misc.";
            toc.Cells[2, 5].Hyperlink = new ExcelHyperLink("'Weighted Dept Salary Portions'!A2", "Weighted Dept Salary Portions");
            toc.Cells[2, 5].Style.Font.UnderLine = true;
            toc.Cells[2, 5].Style.Font.Color.SetColor(Color.Blue);
            toc.Cells[3, 5].Hyperlink = new ExcelHyperLink("'Job Title Weighted Salary Avg.'!A2", "Job Title Weighted Salary Avg.");
            toc.Cells[3, 5].Style.Font.UnderLine = true;
            toc.Cells[3, 5].Style.Font.Color.SetColor(Color.Blue);
            toc.Cells[4, 5].Hyperlink = new ExcelHyperLink("'Department per Job Title Statis'!A2", "Department per Job Title Statistics Summary");
            toc.Cells[4, 5].Style.Font.UnderLine = true;
            toc.Cells[4, 5].Style.Font.Color.SetColor(Color.Blue);
            toc.Cells["E:E"].AutoFitColumns();
        }


        struct uhSpecialties
        {
            public string specialtyCode;
            public string weight;
            public string jobTitle;
            public string department;

            public uhSpecialties(string s, string w, string j, string d)
            {
                specialtyCode = s;
                weight = w;
                jobTitle = j;
                department = d;
            }
        }

        struct tierOneSpecialties
        {
            public string specialtyCode;
            public string averageSalary;
            public string jobTitle;

            public tierOneSpecialties(string s, string ave, string j)
            {
                specialtyCode = s;
                averageSalary = ave;
                jobTitle = j;
            }
        }
    }
}
