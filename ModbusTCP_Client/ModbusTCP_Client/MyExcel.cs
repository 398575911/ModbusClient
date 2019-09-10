using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace ModbusTCP_Client
{
    class MyExcel
    {
        public static string DB_PATH = @"";
        public static BindingList<Brazil_Setup> EmpList = new BindingList<Brazil_Setup>();
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static int lastRow = 0;

        public static void InitializeExcel()
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(DB_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explict cast is not required here
            lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
        }

        public static BindingList<Brazil_Setup> ReadMyExcel()
        {
            EmpList.Clear();
            for (int index = 2; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "J" + index.ToString()).Cells.Value;
                EmpList.Add(new Brazil_Setup
                {
                    Models = MyValues.GetValue(1, 1).ToString(),
                    Model = MyValues.GetValue(1, 2).ToString(),
                    Rated_Pressure = MyValues.GetValue(1, 3).ToString(),
                    Cmin = MyValues.GetValue(1, 4).ToString(),
                    Cmax = MyValues.GetValue(1, 5).ToString(),
                    Powermin = MyValues.GetValue(1, 6).ToString(),
                    Powermax = MyValues.GetValue(1, 7).ToString(),
                    Nozzel_Diameter = MyValues.GetValue(1, 8).ToString(),
                    Nozzel_Coefficient = MyValues.GetValue(1, 9).ToString(),
                    Oil_Type = MyValues.GetValue(1, 10).ToString()
                });
            }
            return EmpList;
        }

        public static void WriteToExcel(Brazil_Setup emp)
        {
            try
            {
                lastRow += 1;
                MySheet.Cells[lastRow, 1] = emp.Models;
                MySheet.Cells[lastRow, 2] = emp.Model;
                MySheet.Cells[lastRow, 3] = emp.Rated_Pressure;
                MySheet.Cells[lastRow, 4] = emp.Cmin;
                MySheet.Cells[lastRow, 5] = emp.Cmax;
                MySheet.Cells[lastRow, 6] = emp.Powermin;
                MySheet.Cells[lastRow, 7] = emp.Powermax;
                MySheet.Cells[lastRow, 8] = emp.Nozzel_Diameter;
                MySheet.Cells[lastRow, 9] = emp.Nozzel_Coefficient;
                MySheet.Cells[lastRow, 10] = emp.Oil_Type;
                EmpList.Add(emp);
                MyBook.Save();
            }
            catch (Exception ex)
            {
                ;
            }

        }

        public static List<Brazil_Setup> FilterEmpList(string searchValue, string searchExpr)
        {
            List<Brazil_Setup> FilteredList = new List<Brazil_Setup>();
            switch (searchValue.ToUpper())
            {
                case "MODELS":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Models.ToLower().Contains(searchExpr));
                    break;
                case "MODEL":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Model.ToLower().Contains(searchExpr));
                    break;
                case "RATED_PRESSURE":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Rated_Pressure.ToLower().Contains(searchExpr));
                    break;
                case "CMIN":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Cmin.ToLower().Contains(searchExpr));
                    break;
                case "CMAX":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Cmax.ToLower().Contains(searchExpr));
                    break;
                case "POWERMIN":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Powermin.ToLower().Contains(searchExpr));
                    break;
                case "POWERMAX":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Powermax.ToLower().Contains(searchExpr));
                    break;
                case "NOZZEL_DIAMETER":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Nozzel_Diameter.ToLower().Contains(searchExpr));
                    break;
                case "NOZZEL_COEFFICIENT":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Nozzel_Coefficient.ToLower().Contains(searchExpr));
                    break;
                case "OIL_TYPE":
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Oil_Type.ToLower().Contains(searchExpr));
                    break;
                default:
                    FilteredList = EmpList.ToList().FindAll(emp => emp.Models.ToLower().Contains(searchExpr));
                    break;
            }
            return FilteredList;
        }

        public static void CloseExcel()
        {
            MyBook.Saved = true;
            MyApp.Quit();
        }

    }
}
