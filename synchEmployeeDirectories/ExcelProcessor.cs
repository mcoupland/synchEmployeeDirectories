using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using System.Drawing;
using System.IO;
using System.Configuration;
using System.Windows;

namespace synchEmployeeDirectories
{
    public class ProgressUpdatedArgs : EventArgs
    {
        private string message;
        public string Message
        {
            get { return message; }
            set { message = value; }
        }

        public ProgressUpdatedArgs(string Message)
        {
            this.Message = Message;
        }
    }

    public delegate void ProgressUpdatedHandler(object sender, ProgressUpdatedArgs e);


    public class ExcelProcessor
    {
        public event ProgressUpdatedHandler ProgressUpdated;
        public event EventHandler ProcessingComplete;

        protected virtual void OnProgressUpdated(string message)
        {
            if (ProgressUpdated != null)
            {
                ProgressUpdated(this, new ProgressUpdatedArgs(message));
            }
        }
        protected virtual void OnProcessingComplete(EventArgs e)
        {
            if (ProcessingComplete != null)
            {
                ProcessingComplete(this, e);
            }
        }


        public static Excel.Application ExcelApplication;
        public static List<Process> ExcelProcesses = new List<Process>();
        private List<Employee> employees = new List<Employee>();
        

        public async void GetUltiproFile()
        {
            ExcelApplication = new Excel.Application() { Visible = false };
            ExcelApplication.UserControl = false;
            ExcelApplication.DisplayAlerts = false;
            Excel.Workbook book = ExcelApplication.Workbooks.Open(ConfigurationManager.AppSettings["ultipro"].ToString());
            Excel.Worksheet sheet = book.Sheets[1];
            Excel.Range range = sheet.UsedRange;

            await Task.Run(() =>
            {     
                try
                {
                    int rowCount = range.Rows.Count;
                    int colCount = range.Columns.Count;

                    employees = new List<Employee>();
                    for (int i = 2; i <= rowCount; i++)
                    {
                        Employee employee = new Employee();
                        employee.Lname = GetValue(range, i, 1);
                        employee.Fname = GetValue(range, i, 2);
                        employee.Initial = GetValue(range, i, 3);
                        employee.Suffix = GetValue(range, i, 4);
                        employee.Nickname = GetValue(range, i, 5);
                        employee.Id = GetValue(range, i, 6);
                        employee.Jobtitle = GetValue(range, i, 7);
                        employee.Division = GetValue(range, i, 8);     
                        employee.Department = GetValue(range, i, 9);  
                        employee.Workphone = GetValue(range, i, 10);
                        employee.Workfax = GetValue(range, i, 11);
                        employee.Workwireless = GetValue(range, i, 12);
                        employee.Employeestatus = GetValue(range, i, 13);
                        employee.Mgrlname = GetValue(range, i, 14);
                        employee.Mgrfname = GetValue(range, i, 15);
                        employee.Mgrmname = GetValue(range, i, 16);
                        employee.Homedepartment = GetValue(range, i, 17);
                        employees.Add(employee);

                        OnProgressUpdated($"{i}/{rowCount-1} Read Employee => {employee.Lname}");
                        Console.WriteLine($"{i} Read Employee {employee.Lname}");
                    }
                }
                finally
                { 
                    
                }
            });
            book.Close();
            ExcelApplication.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(ExcelApplication);
            System.Threading.Thread.Sleep(3000);
            OnProcessingComplete(new EventArgs());
        }

        public async void SaveCeridianFormat()
        {
            ExcelApplication = new Excel.Application() { Visible = false };
            ExcelApplication.UserControl = false;
            ExcelApplication.DisplayAlerts = false;
            Excel.Workbook book = ExcelApplication.Workbooks.Add(Missing.Value);
            Excel.Worksheet sheet = book.Worksheets.get_Item(1);
            Excel.Range range = sheet.UsedRange;


            sheet.Cells[1, 1] = "Last Name";
            sheet.Cells[1, 2] = "First Name";
            sheet.Cells[1, 3] = "Initial";
            sheet.Cells[1, 4] = "Suffix";
            sheet.Cells[1, 5] = "Nickname";
            sheet.Cells[1, 6] = "Employee Id";
            sheet.Cells[1, 7] = "Job Title";
            sheet.Cells[1, 8] = "Division";
            sheet.Cells[1, 9] = "Department";
            sheet.Cells[1, 10] = "Work Phone";
            sheet.Cells[1, 11] = "Work Fax";
            sheet.Cells[1, 12] = "Work Wireless";
            sheet.Cells[1, 13] = "Employee Status Type";
            sheet.Cells[1, 14] = "Mngr. LName";
            sheet.Cells[1, 15] = "Mngr. FName";
            sheet.Cells[1, 16] = "Mngr. MName";
            sheet.Cells[1, 17] = "Home Department (6-digit)";

            ColorConverter colorconverter = new ColorConverter();
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[17]].Interior.Color = ColorTranslator.ToOle((Color) colorconverter.ConvertFromString("#aaaaaa"));
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[17]].Font.Bold = true;

            int row = 2;
            await Task.Run(() =>
            {
                for (int i = 0; i < employees.Count; i++)
                {
                    sheet.Cells[row, 1] = employees[i].Lname;
                    sheet.Cells[row, 2] = employees[i].Fname;
                    sheet.Cells[row, 3] = employees[i].Initial;
                    sheet.Cells[row, 4] = employees[i].Suffix;
                    sheet.Cells[row, 5] = employees[i].Nickname;
                    sheet.Cells[row, 6] = employees[i].Id;
                    sheet.Cells[row, 7] = employees[i].Jobtitle;
                    sheet.Cells[row, 8] = employees[i].Division;
                    sheet.Cells[row, 9] = employees[i].Department;
                    sheet.Cells[row, 10] = employees[i].Workphone;
                    sheet.Cells[row, 11] = employees[i].Workfax;
                    sheet.Cells[row, 12] = employees[i].Workwireless;
                    sheet.Cells[row, 13] = employees[i].Employeestatus;
                    sheet.Cells[row, 14] = employees[i].Mgrlname;
                    sheet.Cells[row, 15] = employees[i].Mgrfname;
                    sheet.Cells[row, 16] = employees[i].Mgrmname;
                    sheet.Cells[row, 17] = employees[i].Homedepartment;
                    row++;
                    OnProgressUpdated($"{i+1}/{employees.Count} Wrote Employee => {employees[i].Lname}");
                    Console.WriteLine($"{i+1} Wrote Employee {employees[i].Lname}");
                }
            });

            sheet.Name = "PhoneDirectory";

            #region Apply Sheet Styles
            range = sheet.UsedRange;
            range.Rows.AutoFit();
            range.Columns.AutoFit();
            range.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            #endregion

            #region Set Page/Printing Options
            var _with1 = sheet.PageSetup;
            _with1.PaperSize = Excel.XlPaperSize.xlPaperLegal;
            _with1.Orientation = Excel.XlPageOrientation.xlLandscape;
            _with1.FitToPagesWide = 1;
            _with1.FitToPagesTall = false;
            _with1.Zoom = false;
            #endregion

            book.SaveAs(
                ConfigurationManager.AppSettings["ceridian"].ToString(),
                Excel.XlFileFormat.xlWorkbookNormal,
                Type.Missing,
                Type.Missing,
                false,
                false,
                Excel.XlSaveAsAccessMode.xlNoChange,
                Excel.XlSaveConflictResolution.xlLocalSessionChanges,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing
            );

            book.Close();
            ExcelApplication.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(ExcelApplication);

            Process.Start(ConfigurationManager.AppSettings["ceridian"].ToString());
        }

        private string GetValue(Excel.Range range, int rowid, int colid)
        {
            var value = range.Cells[rowid, colid].Value2;
            if(value == null)
            {
                value = "";
            }
            else if(value.GetType().Name != "String")
            {
                if(value.GetType().Name == "Int32" && value == -2146826273)
                {
                    value = "0";
                }
                value = Convert.ToString(value).PadLeft(6, '0');

            }
            return value ?? "";
        }


    }

    
}
