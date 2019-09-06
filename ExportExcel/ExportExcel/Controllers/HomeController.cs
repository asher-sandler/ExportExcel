using ExportExcel.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web.Mvc;
using System.Globalization;
using OfficeOpenXml.Style;


namespace ExportExcel.Controllers
{
    public class HomeController : Controller
    {
        // Connect to db
        byonitco_RFID_DBEntities db = new byonitco_RFID_DBEntities();
        // get projectID
        
        int projectID = 0; //  getProjectID();
        // DateTime firstSunday = new DateTime(1753, 1, 7);
        
        public ActionResult Index()
        {
           
            // get Employee List from db

            projectID = getProjectID();

            return View(dataEmpList());
        }

        private List<ProjectsViewModel> dataEmpList()
        {
            return db.TBL_WORK_HOURS_REPORT.Select(x => new ProjectsViewModel
            {
                PROJECT_ID = x.PROJECT_ID,
                EMPLOYEE_NAME = x.EMPLOYEE_NAME,
                WORK_DATE = x.WORK_DATE,
                CLIENT_NAME = x.CLIENT_NAME,
                WORK_DESCRIPTION = x.WORK_DESCRIPTION,
                WORK_HOURS = x.WORK_HOURS,
                HOUR_COST = x.HOUR_COST
                //,DAYOFWEEK = (int)(System.Data.Entity.DbFunctions.DiffDays(firstSunday, x.WORK_DATE) % 7 )
            }).Where(item => item.PROJECT_ID == projectID).ToList();
        }

        //(int)System.Data.Entity.DbFunctions.DiffDays((DateTime?)firstSunday, (DateTime?)expr)) % 7)



        public ActionResult ExportToExcel()
        {


            

            ExcelPackage pkg = new ExcelPackage();
            ExcelWorkbook wb = pkg.Workbook;

            ExcelWorksheet ws = wb.Worksheets.Add("דוח שעות לפי עובדים");
            ws.View.RightToLeft = true;


            // בעבור שמו הגדול
            ws.Cells["A1"].Value =  "בס" +"\"" + "ד";

            
            ws.Cells["C2"].Value = ("דוח שעות לפי עובדים");
            ws.Cells["E2"].Value = "תאריך :";
            ws.Cells["F2"].Value = DateTime.Now.ToString("dd.MM.yyyy HH:mm");

            ws.Cells["G2"].Value = "BYON-IT.COM";


            using (var range = ws.Cells[2, 2, 2, 8])  // Header
            {
                range.Style.Font.Bold = true;
                range.Style.Font.Size = 20;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.DarkSlateGray);
                range.Style.Font.Color.SetColor(Color.White);
            }

            int startRow = 5;

            // list of employee
            IEnumerable<ExportExcel.Models.ProjectsViewModel> emList = db.TBL_WORK_HOURS_REPORT.Select(x => new ExportExcel.Models.ProjectsViewModel
            {
                EMPLOYEE_NAME = x.EMPLOYEE_NAME
            });
            // distinct
            emList = emList.GroupBy(x => x.EMPLOYEE_NAME)
                                  .Select(g => g.First()).OrderBy(c => c.EMPLOYEE_NAME);


            
            string sumGrandTotalFormula = "=";
            string GrandTotalHoursFormula = "=";


            foreach (var empItem in emList)
            {
                
                string sumTotalFormula = "=";
                string sumTotalHoursFormula = "=";

                // emloyee data
                IEnumerable<ExportExcel.Models.ProjectsViewModel> emData = db.TBL_WORK_HOURS_REPORT.Select(x => new ExportExcel.Models.ProjectsViewModel
                {
                    PROJECT_ID = x.PROJECT_ID,
                    EMPLOYEE_NAME = x.EMPLOYEE_NAME,
                    WORK_DATE = x.WORK_DATE,
                    CLIENT_NAME = x.CLIENT_NAME,
                    WORK_DESCRIPTION = x.WORK_DESCRIPTION,
                    WORK_HOURS = x.WORK_HOURS,
                    HOUR_COST = x.HOUR_COST

                }).Where(row => row.EMPLOYEE_NAME == empItem.EMPLOYEE_NAME);

                ws.Cells[string.Format("B{0}", startRow)].Value = empItem.EMPLOYEE_NAME;
                ws.Cells[string.Format("B{0}", startRow)].Style.Font.Size = 20;
                ws.Cells[string.Format("B{0}", startRow)].Style.Font.Bold = true;

                startRow++; startRow++;

                ws.Cells[string.Format("B{0}", startRow)].Value = "תאריך";
                ws.Cells[string.Format("C{0}", startRow)].Value = "יום";
                ws.Cells[string.Format("D{0}", startRow)].Value = "לקוח";
                ws.Cells[string.Format("E{0}", startRow)].Value = "תיאור";
                ws.Cells[string.Format("F{0}", startRow)].Value = "עבודה";
                ws.Cells[string.Format("G{0}", startRow)].Value = "תעריף ₪";
                ws.Cells[string.Format("H{0}", startRow)].Value = "סכום ₪";


                using (var range = ws.Cells[startRow, 2, startRow, 8])  //Address "A1:A5"
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Size = 14;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                    range.Style.Font.Color.SetColor(Color.White);
                }
                startRow++;
                int tableStartRow = startRow;
                foreach (var item in emData)
                {
                    

                    int dayOfWeek = (int)item.WORK_DATE.DayOfWeek;

                    string dayS;
                    switch (dayOfWeek)
                    {
                        case 0: { dayS = "א"; break; }
                        case 1: { dayS = "ב"; break; }
                        case 2: { dayS = "ג"; break; }
                        case 3: { dayS = "ד"; break; }
                        case 4: { dayS = "ה"; break; }
                        case 5: { dayS = "ו"; break; }
                        case 6: { dayS = "ש"; break; }
                        default: { dayS = "Machiah come! End of time."; break; }

                    }

                    ws.Cells[string.Format("B{0}", startRow)].Value = item.WORK_DATE.ToString("dd.MM.yyyy");
                    ws.Cells[string.Format("C{0}", startRow)].Value = dayS;
                    ws.Cells[string.Format("D{0}", startRow)].Value = item.CLIENT_NAME;
                    ws.Cells[string.Format("E{0}", startRow)].Value = item.WORK_DESCRIPTION;
                    ws.Cells[string.Format("F{0}", startRow)].Value = item.WORK_HOURS;
                   
                    ws.Cells[string.Format("G{0}", startRow)].Value =  item.HOUR_COST  ;
                    ws.Cells[string.Format("G{0}", startRow)].Style.Numberformat.Format = "₪#,##0.00";

                    ws.Cells[string.Format("H{0}", startRow)].Formula = "=" + string.Format("G{0}", startRow) + "*" + string.Format("F{0}", startRow);
                    ws.Cells[string.Format("H{0}", startRow)].Style.Numberformat.Format = "₪#,##0.00";


                    sumTotalFormula += string.Format("H{0}", startRow)+"+";
                    sumTotalHoursFormula += string.Format("F{0}", startRow)+"+";

                    sumGrandTotalFormula += string.Format("H{0}", startRow) + "+";
                    GrandTotalHoursFormula += string.Format("F{0}", startRow) + "+";


                    startRow++;

         
                }

                sumTotalFormula += "0";
                sumTotalHoursFormula += "0";
 
                ws.Cells[string.Format("F{0}", startRow)].Formula = sumTotalHoursFormula;
                
   
                ws.Cells[string.Format("H{0}", startRow)].Formula = sumTotalFormula;
                ws.Cells[string.Format("H{0}", startRow)].Style.Numberformat.Format = "₪#,##0.00";
                ws.Cells[string.Format("E{0}", startRow)].Value = "סיכום :";

                using (var range = ws.Cells[startRow, 2, startRow, 8])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Font.Size = 14;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);// SetColor(ColorTranslator.FromHtml(string.Format("azure")));
                    range.Style.Font.Color.SetColor(Color.Orange);
                }

    
                
                startRow++; startRow++;
            }
            startRow++;
            

            sumGrandTotalFormula += "0";
            GrandTotalHoursFormula += "0";


            ws.Cells[string.Format("F{0}", startRow)].Formula = GrandTotalHoursFormula;
            ws.Cells[string.Format("H{0}", startRow)].Formula = sumGrandTotalFormula;
           
            ws.Cells[string.Format("H{0}", startRow)].Style.Numberformat.Format = "₪#,##0.00";
            

            ws.Cells[string.Format("E{0}", startRow)].Value = "סה" + "\"" + "כ" +":";


            using (var range = ws.Cells[startRow, 2, startRow, 8])
            {
                range.Style.Font.Bold = true;
                range.Style.Font.Size = 14;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("Purple")));
                range.Style.Font.Color.SetColor(Color.White);
            }
            ws.Cells["A:AZ"].AutoFitColumns();
            
            
            ws.Column(8).Width = 13;    // Summ


            string outFileName = "דִוּוּחַ" + "_" + getProjectID().ToString() + ".xlsx";
            return File(pkg.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", outFileName);  
        



        }

         private int getProjectID()
        {
            
            int retValue=ToInt32(Request.QueryString["projectID"]);
            if (retValue == 0)
            {
                retValue = 0; //  1425;
               
            }
            return retValue;
        }


         static bool IsNumeric(string s)
         {
             foreach (char c in s)
             {
                 if (!char.IsDigit(c) )
                 {
                     return false;
                 }
             }

             return true;
         }

        static int ToInt32(String value)
        {
            int ret = 0;
            if (!String.IsNullOrEmpty(value))
            {
                if (IsNumeric(value))
                {
                    ret = Int32.Parse(value, CultureInfo.CurrentCulture);
                }
            }
            return ret;
        }

        private string GetDayOfWeekHeb(DateTime workDate)
        {
            int dayOfWeek = (int)workDate.DayOfWeek;

            string dayS; 
            switch (dayOfWeek)
            {
                case 1: { dayS = "א"; break; }
                case 2: { dayS = "ב"; break; }
                case 3: { dayS = "ג"; break; }
                case 4: { dayS = "ד"; break; }
                case 5: { dayS = "ה"; break; }
                case 6: { dayS = "ו"; break; }
                case 7: { dayS = "ש"; break; }
                default: { dayS = "Machiah come! End of time."; break; }
                
            }

            return dayS;
        }
   
    }
}