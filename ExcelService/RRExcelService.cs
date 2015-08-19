using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using ReconRunner.Model;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace ReconRunner.ExcelService
{
    #region Cell
    public enum CellStyle
    {
        Text,
        Currency,
        Bold,
        DarkBlueBold,
        LightBlue,
        Italic
    }

    /// <summary>
    /// This class holds information relevant to populating a cell on an Excel spreadsheet, with information
    /// being column letter, format type, etc..
    /// </summary>
    public class Cell
    {
        private string columnLetter;
        private long rowNumber;
        private CellStyle formatType;
        private string contents;

        public Cell(string cellContents)
        {
            formatType = CellStyle.Text;
            contents = cellContents;
        }

        public Cell(string cellContents, CellStyle cellFormatType)
        {
            formatType = cellFormatType;
            contents = cellContents;
        }

        public string ColumnLetter
        {
            get { return columnLetter; }
            set { columnLetter = value; }
        }

        public long RowNumber
        {
            get { return rowNumber; }
            set { rowNumber = value; }
        }

        public string Contents
        {
            get { return contents; }
            set { contents = value; }
        }

        public CellStyle FormatType
        {
            get { return formatType; }
            set { formatType = value; }
        }

    }
    #endregion Cell

    public class RRExcelService
    {
        #region private properties and class constructors
        static readonly RRExcelService instance = new RRExcelService();
        private Excel.Application ecApp;
        private Excel.Workbook ecWB;
        private Excel.Worksheet ecWS;
        private Excel.Range ecRange;
        private int currRowNumber = 1;
        private object missing = Type.Missing;
        public event ActionStatusEventHandler ActionStatusEvent;

        static RRExcelService()
        {
        }
        public RRExcelService()
        {
        }
        #endregion

        #region properties
        public static RRExcelService Instance
        {
            get { return instance; }
        }
        #endregion

        #region creating rows

        /// <summary>
        /// This will check to see if the excel application object exists.  If it does,
        /// it will quit out of it and set it to null.  Regardless of whether it exists
        /// or not this will also make sure that the 'currRowNumber' is reset back to 1.
        /// </summary>
        public void ResetExcel()
        {
            if (ecApp != null)
            {
                ecApp.Quit();
                ecApp = null;
            }
            currRowNumber = 1;
        }

        /// <summary>
        /// The hashlist will be used to populate a row in the open excel spreadsheet.  
        /// </summary>
        /// <param name="rowValues">A Hashtable whose keys are column letters and values are
        /// the text to put in the corresponding column.</param>
        public void AddRow(Dictionary<string, Cell> rowValues)
        {
            //This will take a hashlist with the keys being column letters, and the values
            //being what text to put in those columns.
            foreach (KeyValuePair<string, Cell> cellEntry in rowValues)
            {
                //The key for each entry is the column letter
                ecRange = ecWS.get_Range(cellEntry.Key + currRowNumber.ToString(), missing);
                ecRange.Value2 = cellEntry.Value.Contents;
                try
                {
                    // This will apply special formatting to the cell.  If the cell has a datatype that is 
                    // unknown it will just be left 'as is'.
                    switch (cellEntry.Value.FormatType)
                    {
                        case CellStyle.Currency:

                            if (cellEntry.Value.Contents == double.MinValue.ToString())
                            {
                                ecRange.Value2 = "n/a";
                            }
                            else
                            {
                                ecRange.Cells.NumberFormat = "$#,##0.#0";
                            }
                            break;
                        case CellStyle.Bold:
                            //Turn the entire cell row bold
                            ecRange.Cells.EntireRow.Font.Bold = true;
                            break;
                        case CellStyle.Italic:
                            //Turn the entire cell row bold
                            ecRange.Cells.EntireRow.Font.Italic = true;
                            break;
                        case CellStyle.DarkBlueBold:
                            //Turn the entire cell row bold
                            ecRange.Cells.EntireRow.Font.Bold = true;
                            ecRange.Cells.Font.Color = ColorTranslator.ToOle(Color.White);
                            ecRange.Interior.Color = ColorTranslator.ToOle(Color.DarkBlue);
                            break;
                        case CellStyle.LightBlue:
                            //Turn the entire cell row bold
                            ecRange.Interior.Color = ColorTranslator.ToOle(Color.LightBlue);
                            break;
                    }
                }
                catch
                {
                }
            }
            ++currRowNumber;
        }

        /// <summary>
        /// Adds a blank row to the current worksheet
        /// </summary>
        public void AddBlankRow()
        {
            Dictionary<string, Cell> excelRow = new Dictionary<string, Cell>();
            excelRow.Add("A", new Cell(" "));
            AddRow(excelRow);
        }

        /// <summary>
        /// Use to add multiple rows at one time.
        /// </summary>
        /// <param name="rows">A SortedList of Dictionary<string, Cell> objects, each representing a row to be added.</param>
        public void AddRows(SortedList<int, Dictionary<string, Cell>> rows)
        {
            foreach (Dictionary<string, Cell> row in rows.Values)
            {
                AddRow(row);
            }
        }

        #endregion creating rows

        /// <summary>
        /// Will save the Excel spreadsheet being created.
        /// </summary>
        /// <param name="filename">The filename to use when saving.</param>
        /// <param name="location">The directory location where the file should be saved.</param>
        /// <returns></returns>
        public string SaveSpreadsheet(string filename)
        {
            //Save the template to the location and filename provided.
            try
            {
                // Do an AutoFitColumn on the spreadsheet before closing it
                AutoFitColumns();
                ecWS.SaveAs(filename, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            }
            catch (Exception e)
            {
                CloseExcel();
                return "Error saving Excel: " + e.Message + "\r\n";
            }
            return "Save successful.";
        }

        /// <summary>
        /// Will perform autofit on the currently open worksheet
        /// </summary>
        /// <returns></returns>
        public void AutoFitColumns()
        {
            ecWS.Cells.EntireColumn.AutoFit();
        }

        public string CloseExcel()
        {
            //Close all Excel objects.
            try
            {
                ecWB.Close(missing, missing, missing);
                ecApp.Quit();
                ecApp = null;
            }
            catch (Exception e)
            {
                return "Error quitting Excel: " + e.Message + "\r\n";
            }

            return "Quit from Excel.\r\n";
        }

        /// <summary>
        /// If there isn't a spreadsheet open yet, it will create it and make this the first worksheet in that
        /// spreadsheet.
        /// 
        /// Otherwise, this will autofit columns on the existing worksheet, create a new worksheet within the 
        /// current open spreadsheet, reset the row counter
        /// and any future calls to AddRow will add to this worksheet until another one is created or the 
        /// spreadsheet is closed.
        /// </summary>
        public void CreateWorksheet(string wsName)
        {
            if (ecApp == null)
            {
                ecApp = new Excel.ApplicationClass();
                //ecApp.ScreenUpdating = false;
                ecWB = ecApp.Workbooks.Add(missing);
                ecWS = (Excel.Worksheet)ecWB.ActiveSheet;
                ecWS.Name = wsName;
            } else 
            {
                AutoFitColumns();
                ecWS = (Excel.Worksheet)ecWB.Worksheets.Add(missing, missing, missing, missing);
                ecWS.Name = wsName;
            }
            
            currRowNumber = 1;        
        }

        /// <summary>
        /// This will remove the worksheet at the index number specified.
        /// </summary>
        /// <param name="wsIndex">An int for the index of the worksheet to be removed.</param>
        public void RemoveWorksheet(int wsIndex)
        {
            ((Excel.Worksheet)ecWB.Sheets[wsIndex]).Delete();
        }

        #region Utilities
        /// <summary>
        /// Use to send non-critical messages.  
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="state"></param>
        /// <param name="message"></param>
        /// <param name="isDetail">Set to false to be seen always, if true may only be seen if user has
        /// requested to see details or in detail logging.</param>
        private void sendActionStatus(object sender, RequestState state, string message, bool isDetail)
        {
            if (ActionStatusEvent != null)
                ActionStatusEvent(sender, new ActionStatusEventArgs(state, message, isDetail));
        }

        #endregion Utilities
    }
}
