//
// EvonsDesigns Excel-Ripper Application for Windows
// The following is copyright 2012 EvonsDesigns
// Author: Joe Evans (evonsdesigns@gmail.com)
//

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using MSExcel = TcKs.MSOffice.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelRipper
{
    public class ExcelManager
    {
        private readonly Form1 _thisForm;

        public ExcelManager(Form1 form)
        {
            _thisForm = form;
        }
        
        /// <summary>
        /// Method that opens a new workbook, then opens each workbook listed (currently only excel 2003), pulls the requested cells out, and closes it.
        /// </summary>
        /// <param name="items">list of filenames</param>
        /// <param name="cellrange">cell range(s)</param>
        /// <param name="outfile">file to save to</param>
        /// <param name="ColumnHeader">Desired column-header for the new excel sheet</param>
        public void Work(List<string> items, string cellrange, string outfile, string ColumnHeader)
        {
            object misValue = Type.Missing;
            Excel.Application myExcelApp = new Excel.Application();
            Excel.Workbooks myExcelWorkbooks = myExcelApp.Workbooks;
            Excel.Workbook outbook = myExcelWorkbooks.Add(misValue);

            myExcelApp.DisplayAlerts = false;
            myExcelApp.ScreenUpdating = false;
            myExcelApp.Visible = false;
            myExcelApp.UserControl = false;
            myExcelApp.Interactive = false;

            try
            {
                Excel.Worksheet ripped = (Excel.Worksheet) outbook.Worksheets[1];

                int cellXPos = 1, cellYPos;

                if(!ColumnHeader.Equals(""))
                {
                    Range columnHeaderRange = ripped.Range["B1"];
                    columnHeaderRange.Value2 = ColumnHeader;
                    cellXPos++;
                }

                for (int i = 0; i < items.Count; i++)
                {
                    string filename = items[i];

                    if (filename.Length > 0)
                    {
                        cellYPos = 2;
                        Range titleRange = ripped.Range["A" + cellXPos];
                        titleRange.Value2 = filename;

                        if (filename.Contains(".xlsx"))
                        {
                            _thisForm.SetStatus("No logic in place for xlsx files. skipping " + Path.GetFileName(filename));
                        }
                        else if (filename.Contains(".xls"))
                        {
                            Excel.Workbook currentExcelWorkbook = myExcelWorkbooks.Open(filename, misValue, misValue, misValue, misValue, misValue,
                                                              misValue, misValue, misValue, misValue, misValue,
                                                               misValue, misValue, misValue, misValue);
                            Excel.Worksheet sheet = currentExcelWorkbook.Worksheets[1]; // could add in logic to have the user specify which worksheet to look in.


                            List<string> cells =
                                new List<string>(cellrange.Split(',').Select(item => item.Trim()).ToArray()); // split the user specified cells into a list of strings.

                            if (cells.Count == 0)
                            {
                                _thisForm.SetStatus("Error: No ranges stated");
                                return;
                            }

                            foreach (string cell in cells)
                            {
                                // go over range
                                if (cell.Contains(":"))
                                {
                                    string[] cellsplit = cell.Split(':');
                                    Excel.Range xRange = sheet.get_Range(cellsplit[0], cellsplit[1]);
                                    object[,] values2DArray = (object[,])xRange.Value2;
                                    foreach (var o in values2DArray)
                                    {
                                        if (o == null)
                                            continue;

                                        string yCol = GetExcelColumnName(cellYPos);
                                        Range newCell = ripped.get_Range(yCol + cellXPos);
                                        newCell.Value2 = o.ToString() ;
                                        cellYPos++;
                                    }
                                }
                                else
                                {
                                    // add one specific cell
                                    Range xRange = sheet.get_Range(cell);
                                    if(xRange.Text != null && !xRange.Text.Equals(""))
                                    {
                                        string yCol = GetExcelColumnName(cellYPos);
                                        Range newCell = ripped.get_Range(yCol + cellXPos);
                                        newCell.Value2 = xRange.Text;
                                        cellYPos++;
                                    }
                                }


                            }
                            cellXPos++;
                            currentExcelWorkbook.Close(false, misValue, misValue);
                        }



                        _thisForm.SetStatus("Successfully opened " + Path.GetFileName(filename));

                        Thread.Sleep(50);
                    }
                }
                _thisForm.SetStatus("Finished ripping. Saving to excel file: " + outfile);
                outbook.Worksheets.Add(ripped);
                outbook.SaveCopyAs(outfile);
                _thisForm.SetStatus("Saved to " +  outfile);
               
                // close any instances  of excel
                outbook.Close(misValue, misValue, misValue);
                myExcelWorkbooks.Close();
                myExcelApp.Quit();
                Release(outbook);
                Release(myExcelWorkbooks);
                Release(myExcelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch(Exception ex)
            {
                _thisForm.Text = "Error occurred: "+ex.Message;
                // close any instances  of excel
                outbook.Close(misValue, misValue, misValue);
                myExcelWorkbooks.Close();
                myExcelApp.Quit();
                Release(outbook);
                Release(myExcelWorkbooks);
                Release(myExcelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void Release(object obj)
        {
            // Errors are ignored per Microsoft's suggestion for this type of function:
            // http://support.microsoft.com/default.aspx/kb/317109
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
            }
            catch
            {
            }
        }

        /// <summary>
        /// return the column name from a given number value
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString(CultureInfo.InvariantCulture) + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }


    }
}
