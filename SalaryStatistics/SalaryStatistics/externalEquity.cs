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

            ExcelWorksheet weightedDeptAvg = excelFile.Workbook.Worksheets.Add("Weighted Dept Salary Avg");
            weightedDeptAvg.Cells[1, 1].Value = "Department";
            weightedDeptAvg.Cells[1, 2].Value = "UH Average";
            weightedDeptAvg.Cells[1, 3].Value = "Tier 1 Weighted Average";
            weightedDeptAvg.Cells[1, 4].Value = "Compression";

            currentRow = 2;
            List<string> Departments = getDepartments("Departments");
            foreach (string deptName in Departments)
            {    
                weightedDeptAvg.Cells[currentRow, 1].Value = deptName;
                weightedDeptAvg.Cells[currentRow, 2].Formula = "'" + deptName + "'" + "!E2";

                int firstRow = 0;
                int counter = 0;
                bool toBreak = false;
                bool wasFound = false;
                for (int row = 2; row <= weightedDeptPortions.Dimension.End.Row; row++)
                {
                   firstRow = row;

                   while (weightedDeptPortions.Cells["B"+row].Value.ToString().ToUpper().Equals(deptName.ToString().ToUpper()))
                   {
                       counter++;

                       if (!weightedDeptPortions.Cells["B" + (row + 1)].Value.ToString().Equals(deptName))
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
                    weightedDeptAvg.Cells[currentRow, 3].Formula = "SUM('Weighted Dept Salary Portions'!D" + firstRow + ":" + "'Weighted Dept Salary Portions'!D" + (firstRow + counter - 1) + ")";
                }
                else
                {
                    weightedDeptAvg.Cells[currentRow, 3].Value = "Department not found";
                }
                weightedDeptAvg.Cells[currentRow, 4].Formula = "C" + currentRow + "/" + "B" + currentRow;
                currentRow++;
            }

            weightedDeptAvg.Cells["B:C"].Style.Numberformat.Format = "$###,###,##0";
            weightedDeptAvg.Cells["A:Z"].AutoFitColumns();

            ExcelWorksheet toc = excelFile.Workbook.Worksheets["Table Of Contents"];

            toc.Cells[1, 5].Value = "Misc.";
            toc.Cells[2, 5].Hyperlink = new ExcelHyperLink("'Weighted Dept Salary Portions'!A2", "Weighted Dept Salary Portions");
            toc.Cells[2, 5].Style.Font.UnderLine = true;
            toc.Cells[2, 5].Style.Font.Color.SetColor(Color.Blue);
            toc.Cells[3, 5].Hyperlink = new ExcelHyperLink("'Weighted Dept Salary Avg'!A2", "Weighted Dept Salary Avg");
            toc.Cells[3, 5].Style.Font.UnderLine = true;
            toc.Cells[3, 5].Style.Font.Color.SetColor(Color.Blue);
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
