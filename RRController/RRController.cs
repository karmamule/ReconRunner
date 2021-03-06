using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using ReconRunner;
using ReconRunner.ExcelService;
using ReconRunner.Model;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace ReconRunner.Controller
{
    public sealed class RRController
    {
        #region Properties and constructors
        static readonly RRController instance = new RRController();

        private RRDataService rrDataService = RRDataService.Instance;
        private RRExcelService rrExcelService = RRExcelService.Instance;
        private string firstQueryLabel;
        private string secondQueryLabel;
        private List<ReconReportAndData> reconReportsAndData = new List<ReconReportAndData>();

        public event ActionStatusEventHandler ActionStatusEvent;
        string pipePlaceholderRegex = @"\|(\w+)\|";

        // Each 2-query recon is divided into three sections:
        //      Rows found in the first query without a matching row in the second
        //      Rows found in the second query without a matching row in the first
        //      Rows found in both queries that have one or more columns that do not match
        // A single query recon has just one section, ProblemRows
        enum reportSection
        {
            FirstQueryRowWithoutMatch,
            SecondQueryRowWithoutMatch,
            DataDifferences,
            ProblemRows
        }
        private Recons recons = new Recons();
        public Recons Recons
        {
            get { return recons; }
            set {
                    recons = value;
                    if (!IsDataValid)
                        sendActionStatus(this, RequestState.DataInvalid, "Data is not currently valid. Use GetValidationErrors() for details", false);
                    else
                        sendActionStatus(this, RequestState.DataValid, "Sources and recons are valid.", false);
                }
        }
        private Sources sources = new Sources();
        public Sources Sources
        {
            get { return sources; }
            set {
                    sources = value;
                    if (!IsDataValid)
                        sendActionStatus(this, RequestState.DataInvalid, "Data is not currently valid. Use GetValidationErrors() for details", false);
                    else
                        sendActionStatus(this, RequestState.DataValid, "Sources and recons are valid.", false);
            }
        }

        public bool IsDataValid
        {
            get
            {
                if (validateRecons().Count > 0 || validateSources().Count > 0)
                    return false;
                else
                    return true;
            }
        }

        public List<ReconReport> ReconReports
        {
            get { return recons.ReconReports; }
            set { recons.ReconReports = value; }
        }

        // Use to translate index to column letter when # columns is variable
        string columnLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";  // *** ASSUMPTION: We'll never go past Z in columns
        // These will hold DataRow objects, with the string index key for each being built from its identifying columns
        //DataTable firstQueryData = new DataTable();
        //DataTable secondQueryData = new DataTable();
        Dictionary<string, DataRow> q1Rows;
        Dictionary<string, DataRow> q2Rows;
        Dictionary<string, Cell> excelRow;

        private RRController()
        {
            //rrDataService.ActionStatusEvent += new ActionStatusEventHandler(relayActionStatusEvent);
            //rrExcelService.ActionStatusEvent += new ActionStatusEventHandler(relayActionStatusEvent);
        }

        public static RRController Instance
        {
            get
            {
                return instance;
            }
        }

        #endregion Properties and constructors

        #region recons object serialization
        /// <summary>
        /// Creates a sample XML file populated with XML for the recon XML file
        /// that specifies recon reports and query substituation variable values to be run
        /// </summary>
        /// <param name="fileName"></param>
        public void CreateSampleXMLReconReportFile(string fileName)
        {
            getNewFileService().WriteSampleReconsToXMLFile(fileName);
            sendActionStatus(this, RequestState.Succeeded, "Created sample recons XML file.", false);
        }

        /// <summary>
        /// Creates a sample XML file populated with XML for the sources XML file
        /// that holds database connection templates, connection strings, and queries
        /// to be used in making recon reports
        /// </summary>
        /// <param name="fileName"></param>
        public void CreateSampleXMLSourcesFile(string fileName)
        {
            getNewFileService().WriteSampleSourcesToXMLFile(fileName);
            sendActionStatus(this, RequestState.Succeeded, "Created sample sources XML file.", false);
        }

        public void WriteReconsToXmlFile(string fileName)
        {
            getNewFileService().WriteReconsToXMLFile(recons, fileName);
        }

        public void WriteSourcesToXmlFile(string fileName)
        {
            getNewFileService().WriteSourcesToXMLFile(sources, fileName);
        }

        /// <summary>
        /// Read in information for all the recons to be done from the specified xml file.  Note that 
        /// each recon query will need a corresponding entry in App.Config that specifies the actual
        /// SQL statement and connection string to be used for each query.
        /// </summary>
        /// <param name="fileName"></param>
        public void ReadReconsFromXMLFile(string fileName)
        {
            try
            {
                Recons = getNewFileService().ReadReconsFromXMLFile(fileName);
                sendActionStatus(this, RequestState.Succeeded, "Read recons from XML file.", false);
            }
            catch (Exception e)
            {
                Recons = null;
                sendActionStatus(this, RequestState.Error, string.Format("Failed to read recons from XML file: {0}.", GetFullErrorMessage(e)), false);            
            }
        }

        /* *****TO DO MONDAY 14 JULY: Modify logic creating datatable(s) for query(ies) to use new generic dataservice that reads
         * Sources and other objects to make appropriate connections and recons to execute SQL queries. 
         * 
         */

        /// <summary>
        /// Read in information for all the sources (data source templates, connection strings, and queries) 
        /// needed for the recons to be run.
        /// </summary>
        /// <param name="fileName"></param>
        public void ReadSourcesFromXMLFile(string fileName)
        {
            try
            {
                Sources = getNewFileService().ReadSourcesFromXMLFile(fileName);
                sendActionStatus(this, RequestState.Succeeded, "Read sources from XML file.", false);
            }
            catch (Exception e)
            {
                Sources = null;
                sendActionStatus(this, RequestState.Error, string.Format("Failed to read sources from XML file: {0}.", GetFullErrorMessage(e)), false);
            }
        }

        #endregion recons object serialization

        /// <summary>
        /// Once the recons object has been populated, this may be run to
        /// actually execute each of the recons and write their results to a
        /// spreadsheet.  Each recon's output will be written to its own tab
        /// on the spreadsheet.
        /// </summary>
        /// <returns>A string with information on the run (success/failure/reports run/etc.)</returns>
        public void RunRecons(string excelFileName)
        {
            if (sources.Queries.Count == 0 || recons.ReconReports.Count == 0) 
                throw new Exception("Either sources or recons not specified");
            rrDataService.Sources = sources;
            rrDataService.Recons = recons;
            List<DataTable> reconData = new List<DataTable>();

            try
            {
                rrDataService.OpenConnections(recons);
            }
            catch (Exception ex)
            {
                rrDataService.CloseConnections();
                throw new Exception("Error while trying to open connections: " + GetFullErrorMessage(ex));
            }

            try
            {
                List<ReconReportAndData> reportsAndData = getReconData();
                createReconTabs(reportsAndData);
            }
            catch(Exception e)
            {
                Exception originalException = e;
                try
                {
                    rrDataService.CloseConnections();
                    rrExcelService.CloseExcel();
                }
                catch (Exception newEx)
                {
                    string fullErrorMessage = string.Format("While cleaning up after error {0} ran into error {1}. Connections and/or Excel may not have closed successfully.", GetFullErrorMessage(originalException), GetFullErrorMessage(newEx));
                    sendActionStatus(this, RequestState.FatalError, fullErrorMessage, false);
                }
                sendActionStatus(this, RequestState.FatalError, string.Format("Error while running recon: {0}. Successfully closed connections and Excel.", GetFullErrorMessage(e)), false);
            }
            rrDataService.CloseConnections();
            sendActionStatus(this, RequestState.Information, "Closed all connections.", false);

            // Save the spreadsheet and close the Excel process.
            rrExcelService.SaveSpreadsheet(excelFileName);
            rrExcelService.CloseExcel();
            sendActionStatus(this, RequestState.Succeeded, "Saved spreadsheet, closed Excel, Done.", false);
        }

        /// <summary>
        /// Modified to now run any recon reports' queries in parallel when the report is set to permit
        /// it. Any reports that are NOT marked to allow parallel execution will then have their queries
        /// run sequentially.
        /// </summary>
        /// <returns></returns>
        private List<ReconReportAndData> getReconData()
        {
            reconReportsAndData = new List<ReconReportAndData>();
            Parallel.ForEach(recons.ReconReports.FindAll(rr => rr.ParallelOk), getReportData);
            recons.ReconReports.FindAll(rr => !rr.ParallelOk).ForEach(getReportData);
            return reconReportsAndData;
        }

        private void getReportData(ReconReport recon)
        {
                var reconAndData = new ReconReportAndData(recon);
                sendActionStatus(this, RequestState.Information, string.Format("Getting data for recon {0}.", recon.Name), false);
                var reconData = rrDataService.GetReconData(recon);
                reconAndData.FirstQueryData = reconData[0];
                if (recon.SecondQueryName != "")
                    reconAndData.SecondQueryData = reconData[1];
                reconReportsAndData.Add(reconAndData);
                sendActionStatus(this, RequestState.Information, string.Format("Finished getting data for recon {0}.", recon.Name), false);
        }

        private void createReconTabs(List<ReconReportAndData> reportsAndData)
        {
            reportsAndData.ForEach(reconAndData =>
            {
            sendActionStatus(this, RequestState.Information, string.Format("Creating Excel tab {0} for recon {1}.", reconAndData.ReconReport.TabLabel, reconAndData.ReconReport.Name), false);
            createReconTab(reconAndData.ReconReport, reconAndData.FirstQueryData, reconAndData.SecondQueryData);
            sendActionStatus(this, RequestState.Succeeded, string.Format("Finished Excel tab {0} for recon {0}.", reconAndData.ReconReport.TabLabel, reconAndData.ReconReport.Name), false);
            });
        }

        /// <summary>
        /// Take one or two queries' data and create the corresponding tab in the spreadsheet.  For recon involving 
        /// two queries it will correlate the two rows and create a tab with header plus 3 sections.  If a single query
        /// report (recon has null second query) then it will just create a tab with header and 1 section.
        /// </summary>
        /// <param name="recon"></param>
        /// <param name="firstQueryData"></param>
        /// <param name="secondQueryData">May pass empty datatable for single query recons</param>
        private void createReconTab(ReconReport recon, DataTable firstQueryData, DataTable secondQueryData)
        {
            int numQueries = recon.SecondQueryName == "" ? 1 : 2;
            if (recon.SecondQueryName != "")
                populateIndexedRowCollections(recon, firstQueryData, secondQueryData);
            // Create a new worksheet within our spreadsheet to hold the results
            rrExcelService.CreateWorksheet(recon.TabLabel);
            // Write the title rows
            writeWorkSheetHeaderRows(recon);
            // The next two sections are only needed for two-query recons
            if (recon.SecondQueryName != "")
            {
                // Write the first section showing rows in the first query without a match in the second query
                writeOrphanFirstQuerySection(recon);
                rrExcelService.AddBlankRow();
                rrExcelService.AddBlankRow();
                // Write the second section that shows rows in the second query without a match in the first query
                writeOrphanSecondQuerySection(recon);
                rrExcelService.AddBlankRow();
                rrExcelService.AddBlankRow();
            }
            // Create the section of the report that either shows all the rows returned by a single query recon, or
            // is the comparison section of a 2-query recon that shows any pairs of records that have differing values
            // in one or more columns that are supposed to match
            writeDataProblems(recon, firstQueryData, secondQueryData);
        }

        /// <summary>
        /// Prepares the data for writing to spreadsheet for 2-query recons by creating a key for each datarow
        /// that is created by concatenating the values for any columns designated as identity columns.  Not necessary
        /// for single query recons.
        /// </summary>
        /// <param name="recon"></param>
        /// <param name="q1Data"></param>
        /// <param name="q2Data"></param>
        private void populateIndexedRowCollections(ReconReport recon, DataTable q1Data, DataTable q2Data)
        {
            string rowKey;
            q1Rows = new Dictionary<string, DataRow>();
            q2Rows = new Dictionary<string, DataRow>();

            try
            {
                foreach (DataRow row in q1Data.Rows)
                {
                    rowKey = "";
                    foreach (QueryColumn column in recon.Columns)
                    {
                        if (column.IdentifyingColumn)
                        {
                            rowKey += row[column.FirstQueryColName].ToString();
                        }
                    }
                    q1Rows.Add(rowKey, row);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error (most likely identifying columns not sufficient to guarantee uniqueness) while preparing first query data: " + GetFullErrorMessage(ex));
            }

            try
            {
                foreach (DataRow row in q2Data.Rows)
                {
                    rowKey = "";
                    foreach (QueryColumn column in recon.Columns)
                    {
                        if (column.IdentifyingColumn)
                        {
                            rowKey += row[column.SecondQueryColName].ToString();
                        }
                    }
                    q2Rows.Add(rowKey, row);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error (most likely identifying columns not sufficient to guarantee uniqueness) while preparing second query data: " + GetFullErrorMessage(ex));
            }
        }

        /// <summary>
        /// Creates the section of the recon that shows rows in the first query without a match found
        /// in the second query.
        /// </summary>
        /// <param name="recon"></param>
        private void writeOrphanFirstQuerySection(ReconReport recon)
        {
            CellStyle currStyle = CellStyle.LightBlue;
            int orphanCounter = 0;
            string counterText;
            setReconQueryLabels(recon);
            // Start the section with a label
            excelRow = new Dictionary<string, Cell>();
            excelRow.Add("A", new Cell("Rows in ", CellStyle.Bold));
            excelRow.Add("B", new Cell(firstQueryLabel, CellStyle.Bold));
            excelRow.Add("C", new Cell("Without a Match in", CellStyle.Bold));
            excelRow.Add("D", new Cell(secondQueryLabel, CellStyle.Bold));
            rrExcelService.AddRow(excelRow);
            // Now write out the common columns at the left of each section
            writeCommonHeaders(2, recon.Columns, true, true);
            // There are no additional columns for this section, so now write out the identifying and 'always display' columns
            // in the first query that don't have a match in the second query.
            foreach (string rowKey in q1Rows.Keys)
            {
                if (!q2Rows.ContainsKey(rowKey))
                {
                    // Toggle background color of the row being written
                    currStyle = toggleStyle(currStyle);
                    // We have a row in the first query without a match in the second
                    writeRowCommonColumns(q1Rows[rowKey], recon.Columns, reportSection.FirstQueryRowWithoutMatch, currStyle);
                    ++orphanCounter;
                }
            }
            excelRow.Clear();
            if (orphanCounter == 1)
            {
                counterText = orphanCounter.ToString() + " orphan found.";
            }
            else
            {
                counterText = orphanCounter.ToString() + " orphans found.";
            }
            excelRow.Add("A", new Cell(counterText));
            rrExcelService.AddRow(excelRow);
        }

        /// <summary>
        /// Creates the section of the recon that shows rows in the second query without a match found
        /// in the first query.
        /// </summary>
        /// <param name="recon"></param>
        private void writeOrphanSecondQuerySection(ReconReport recon)
        {
            CellStyle currStyle = CellStyle.LightBlue;
            int orphanCounter = 0;
            string counterText;
            setReconQueryLabels(recon);
            // Start the section with a label
            excelRow = new Dictionary<string, Cell>();
            excelRow.Add("A", new Cell("Rows in ", CellStyle.Bold));
            excelRow.Add("B", new Cell(secondQueryLabel, CellStyle.Bold));
            excelRow.Add("C", new Cell("Without a Match in", CellStyle.Bold));
            excelRow.Add("D", new Cell(firstQueryLabel, CellStyle.Bold));
            rrExcelService.AddRow(excelRow);
            // Now write out the common columns at the left of each section

            writeCommonHeaders(2, recon.Columns, true, true);
            // There are no additional columns for this section, so now write out the identifying and 'always display' columns
            // in the first query that don't have a match in the second query.
            foreach (string rowKey in q2Rows.Keys)
            {
                if (!q1Rows.ContainsKey(rowKey))
                {
                    // Toggle background color of the row being written
                    currStyle = toggleStyle(currStyle);
                    // We have a row in the first query without a match in the second
                    writeRowCommonColumns(q2Rows[rowKey], recon.Columns, reportSection.SecondQueryRowWithoutMatch, currStyle);
                    ++orphanCounter;
                }
            }
            excelRow.Clear();
            if (orphanCounter == 1)
            {
                counterText = orphanCounter.ToString() + " orphan found.";
            }
            else
            {
                counterText = orphanCounter.ToString() + " orphans found.";
            }
            excelRow.Add("A", new Cell(counterText));
            rrExcelService.AddRow(excelRow);
        }

        /// <summary>
        /// If receives CellStyle Text, returns CellStyle.LightBlue.  If receives style LightBlue
        /// (or any other style) it returns CellStyle.Text.
        /// </summary>
        /// <param name="currStyle">CellStyle.Text or CellStyle.LightBlue</param>
        /// <returns></returns>
        private CellStyle toggleStyle(CellStyle currStyle)
        {
            if (currStyle == CellStyle.Text)
            {
                return CellStyle.LightBlue;
            }
            else
            {
                return CellStyle.Text;
            }      
        }

        /// <summary>
        /// If a single query recon will show all the rows returned by the query.  If a two query recon
        /// then will show all the rows matched between the two queries that have data differences in the
        /// columns that are marked "ShouldMatch"
        /// </summary>
        /// <param name="recon"></param>
        /// <param name="firstQueryData"></param>
        /// <param name="secondQueryData"></param>
        private void writeDataProblems(ReconReport recon, DataTable firstQueryData, DataTable secondQueryData)
        {
            CellStyle currStyle = CellStyle.LightBlue;
            int currColumnIndex;
            int numDifferencesFound = 0;
            int numQueries = recon.SecondQueryName == "" ? 1 : 2;
            string counterText;
            setReconQueryLabels(recon);
            // Start the section with a label if 2-query recon
            if (recon.SecondQueryName != "")
            {
                excelRow = new Dictionary<string, Cell>();
                excelRow.Add("A", new Cell("Rows in ", CellStyle.Bold));
                excelRow.Add("B", new Cell(firstQueryLabel, CellStyle.Bold));
                excelRow.Add("C", new Cell("that have one or more", CellStyle.Bold));
                excelRow.Add("D", new Cell("differences compared to", CellStyle.Bold));
                excelRow.Add("E", new Cell(secondQueryLabel, CellStyle.Bold));
                rrExcelService.AddRow(excelRow);
            }
            // Now write out the common columns at the left of each section
            // If a 1 query recon then we don't need extra columns and should add the header row as is
            currColumnIndex = writeCommonHeaders(numQueries, recon.Columns, recon.SecondQueryName == "", false);
            if (recon.SecondQueryName != "")
            {
                // We have to add three more columns: one to list the item the differences was found in, and then
                // two columns to show the two values found.
                excelRow.Add(columnLetters[currColumnIndex].ToString(), new Cell("Column", CellStyle.DarkBlueBold));
                excelRow.Add(columnLetters[currColumnIndex + 1].ToString(), new Cell("1st Query Value", CellStyle.DarkBlueBold));
                excelRow.Add(columnLetters[currColumnIndex + 2].ToString(), new Cell("2nd Query Value", CellStyle.DarkBlueBold));
                rrExcelService.AddRow(excelRow);
            }
            // Now go through the queries and write one row per row returned in the case of 1-query recons or one row
            // per difference found in the case of 2-query recons
            if (recon.SecondQueryName == "")
            {
                // Write out rows for 1-query recons
                if (firstQueryData.Rows.Count == 0)
                {
                    excelRow = new Dictionary<string, Cell>();
                    excelRow.Add("A", new Cell("No rows returned", CellStyle.Bold));
                    rrExcelService.AddRow(excelRow);
                }
                else
                {
                    foreach (DataRow row in firstQueryData.Rows)
                    {
                        excelRow = new Dictionary<string, Cell>();
                        currStyle = toggleStyle(currStyle);
                        // Always show all columns on a 1-query recon
                        for (int n = 0; n < recon.Columns.Count; n++)
                        {
                            excelRow.Add(columnLetters[n].ToString(), new Cell(row[recon.Columns[n].FirstQueryColName].ToString(), currStyle));
                        }
                        rrExcelService.AddRow(excelRow);
                    }
                    // Write a summary row with number of rows returned
                    counterText = string.Format("{0} {1} returned.", firstQueryData.Rows.Count, firstQueryData.Rows.Count == 1 ? "row" : "rows");
                    excelRow.Clear();
                    excelRow.Add("A", new Cell(counterText));
                    rrExcelService.AddRow(excelRow);                }
            }
            else
            {   // Write out rows for 2-query recons
                foreach (string rowKey in q1Rows.Keys)
                {
                    if (q2Rows.ContainsKey(rowKey))
                    {
                        // Toggle background color of the row being written
                        currStyle = toggleStyle(currStyle);
                        // Get the two matched rows and see if any of their columns that should match
                        // actually have any differences that need to be reported.
                        foreach (QueryColumn column in recon.Columns)
                        {
                            if (column.CheckDataMatch)
                            {
                                if (!valuesMatch(q1Rows[rowKey][column.FirstQueryColName], q2Rows[rowKey][column.SecondQueryColName], column))
                                {
                                    // We have a mismatch that needs to be reported
                                    // First, write the identifying and 'Always Display' columns, and get the column index we're at
                                    currColumnIndex = writeDiffRowCommonColumns(q1Rows[rowKey], q2Rows[rowKey], recon.Columns, currStyle);
                                    // Now write out the three columns regarding the difference found: 
                                    //      column name, first query value, and second query value
                                    excelRow.Add(columnLetters[currColumnIndex].ToString(), new Cell(column.Label, currStyle));
                                    excelRow.Add(columnLetters[currColumnIndex + 1].ToString(), new Cell(q1Rows[rowKey][column.FirstQueryColName].ToString(), currStyle));
                                    excelRow.Add(columnLetters[currColumnIndex + 2].ToString(), new Cell(q2Rows[rowKey][column.SecondQueryColName].ToString(), currStyle));
                                    // Finally, add this row to the recon report
                                    rrExcelService.AddRow(excelRow);
                                    ++numDifferencesFound;
                                }
                            }
                        }
                    }
                }
                // If no differences were found, write one row to that effect
                if (numDifferencesFound == 0)
                {
                    excelRow = new Dictionary<string, Cell>();
                    excelRow.Add("A", new Cell("No differences found", CellStyle.Bold));
                    rrExcelService.AddRow(excelRow);
                }
                else
                {
                    excelRow.Clear();
                    if (numDifferencesFound == 1)
                    {
                        counterText = numDifferencesFound.ToString() + " difference found.";
                    }
                    else
                    {
                        counterText = numDifferencesFound.ToString() + " differences found.";
                    }
                    excelRow.Add("A", new Cell(counterText));
                    rrExcelService.AddRow(excelRow);
                }
            }
        }

        /// <summary>
        /// Will return true if the two values match.
        /// </summary>
        /// <param name="value1"></param>
        /// <param name="value2"></param>
        /// <param name="column">The column these values are for</param>
        /// <returns></returns>
        private bool valuesMatch(object value1, object value2, QueryColumn column)
        {
            // If only one value is null, return false
            if ((value1 == System.DBNull.Value && value2 != System.DBNull.Value) || (value1 != System.DBNull.Value && value2 == System.DBNull.Value))
            {
                return false;
            }

            // If both values are null, return true
            if (value1 == System.DBNull.Value && value2 == System.DBNull.Value)
            {
                return true;
            }

            // Neither value is null, so compare them based on what type they should be
            // If either one fails to parse as the proper value type, then return false
            switch (column.Type) 
            {
                case ColumnType.date:
                    DateTime q1DateTime;
                    DateTime q2DateTime;
                    if (!DateTime.TryParse(value1.ToString(), out q1DateTime) || 
                        !DateTime.TryParse(value2.ToString(), out q2DateTime))
                    {
                        return false;
                    }
                    if (q1DateTime == q2DateTime) {
                        return true;
                    }
                    break;
                case ColumnType.number:
                // Columns are considered a match if their difference is less than or equal to any tolerance threshold provided
                    decimal q1Decimal;
                    decimal q2Decimal;
                    if (!decimal.TryParse(value1.ToString(), out q1Decimal) ||
                        !decimal.TryParse(value2.ToString(), out q2Decimal))
                    {
                        return false;
                    }
                    if (q1Decimal == q2Decimal || (Math.Abs(q1Decimal - q2Decimal) <= column.Tolerance)) {
                        return true;
                    }
                    break;
                case ColumnType.text:
                    if (value1.ToString() == value2.ToString()) {
                        return true;
                    }
                    break;
            }
            return false;
        }
        
        /// <summary>
        /// Write out the columns we must display for this row.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="recon"></param>
        /// <param name="formatType">What CellFormatType to use for the cell</param>
        private void writeRowCommonColumns(DataRow row, List<QueryColumn> columns, reportSection section, CellStyle formatType)
        {
            QueryColumn queryColumn;
            excelRow = new Dictionary<string, Cell>();
            string columnName;
            int currColumnIndex = 0;

            // Write out identifying columns
            for (int i = 0; i < columns.Count; i++)
            {
                queryColumn = columns[i];
                if (queryColumn.IdentifyingColumn)
                {
                    if (section == reportSection.FirstQueryRowWithoutMatch)
                    {
                        columnName = queryColumn.FirstQueryColName;
                    }
                    else
                    {
                        columnName = queryColumn.SecondQueryColName;
                    }
                    excelRow.Add(columnLetters[currColumnIndex].ToString(), new Cell(row[columnName].ToString(), formatType));
                    ++currColumnIndex;
                }
            }
            // Write out columns marked AlwaysDisplay (that aren't also identifying columns).  
            for (int i = 0; i < columns.Count; i++)
            {
                queryColumn = columns[i];
                if ((queryColumn.AlwaysDisplay || queryColumn.CheckDataMatch) && !queryColumn.IdentifyingColumn)
                {
                    if (section == reportSection.FirstQueryRowWithoutMatch)
                    {
                        columnName = queryColumn.FirstQueryColName;
                    }
                    else
                    {
                        columnName = queryColumn.SecondQueryColName;
                    }
                    // Some columns marked 'AlwaysDisplay' may not exist on one of the two queries
                    if (columnName != null)
                    {
                        excelRow.Add(columnLetters[currColumnIndex].ToString(), new Cell(row[columnName].ToString(), formatType));
                    }
                    else
                    {
                        excelRow.Add(columnLetters[currColumnIndex].ToString(), new Cell("", formatType));
                    }
                    ++currColumnIndex;
                }
            }
            rrExcelService.AddRow(excelRow);
        }

        /// <summary>
        /// Write out the columns we must display for this row.
        /// </summary>
        /// <param name="row"></param>
        /// <param name="recon"></param>
        /// <param name="formatType">What CellFormatType to use for the cell</param>
        private int writeDiffRowCommonColumns(DataRow q1Row, DataRow q2Row, List<QueryColumn> columns, CellStyle formatType)
        {
            QueryColumn queryColumn;
            excelRow = new Dictionary<string, Cell>();
            string columnValue;
            int currColumnIndex = 0;

            // Write out identifying columns
            for (int i = 0; i < columns.Count; i++)
            {
                queryColumn = columns[i];
                if (queryColumn.IdentifyingColumn)
                {
                    excelRow.Add(columnLetters[currColumnIndex].ToString(), new Cell(q1Row[queryColumn.FirstQueryColName].ToString(), formatType));
                    ++currColumnIndex;
                }
            }
            // Write out columns marked AlwaysDisplay (that aren't also identifying columns).  
            for (int i = 0; i < columns.Count; i++)
            {
                queryColumn = columns[i];
                if (queryColumn.AlwaysDisplay && !queryColumn.IdentifyingColumn)
                {
                    // Some columns marked 'AlwaysDisplay' may not exist on one of the two queries.
                    // Use value on first query if it exists on either the first query or both, otherwise use
                    // the value in the second query.
                    if (queryColumn.FirstQueryColName != null)
                    {
                        columnValue = q1Row[queryColumn.FirstQueryColName].ToString();
                    }
                    else
                    {
                        columnValue = q2Row[queryColumn.SecondQueryColName].ToString();
                    }
                    excelRow.Add(columnLetters[currColumnIndex].ToString(), new Cell(columnValue, formatType));
                    ++currColumnIndex;
                }
            }
            return currColumnIndex;
        }

        /// <summary>
        /// This writes out header columns for all columns of the recon that will be displayed regardless
        /// of what section of the recon we're showing.  If a single query recon write all columns, else
        /// only display columns as needed
        /// </summary>
        /// <param name="numQueries">Number of queries in the recon.  Either 1 or 2.</param>
        /// <param name="columns">Collection of columns to beused</param>
        /// <param name="addRow">set to true to add row to spreadsheet and move to next row</param>
        /// <param name="showShouldMatchColumns">Set to true to show all ShouldMatch column names in the header. 
        /// (This will be true for the "orphan" report sections, and false for the data differences section.</param>
        /// <returns>
        /// An integer with the value of the currColumnIndex, in case the caller needs to tack additional
        /// columns after these before adding the row to the worksheet.
        /// </returns>
        private int writeCommonHeaders(int numQueries, List<QueryColumn> columns, bool addRow, bool showShouldMatchColumns)
        {
            int currColumnIndex = 0;
            QueryColumn queryColumn;
            excelRow = new Dictionary<string, Cell>();
            string columnAttribute;
            
            for (int i = 0; i < columns.Count; i++)
            {
                queryColumn = columns[i];
                if (queryColumn.IdentifyingColumn || numQueries == 1)
                {
                    var idAnnotation = numQueries == 1 ? string.Empty : " (Id)";
                    excelRow.Add(columnLetters[currColumnIndex].ToString(), new Cell(queryColumn.Label + idAnnotation, CellStyle.DarkBlueBold));
                    ++currColumnIndex;
                }
            }
            // Write out columns marked AlwaysDisplay.  
            // Yes this is a second iteration through the same set of columns, but we want to make sure all
            // identifying columns precede all columns marked 'AlwaysDisplay' even if they are not in proper order
            // in the collection
            if (numQueries == 2) 
            for (int i = 0; i < columns.Count; i++)
            {
                queryColumn = columns[i];
                if ((queryColumn.AlwaysDisplay || (queryColumn.CheckDataMatch && showShouldMatchColumns)) && !queryColumn.IdentifyingColumn)
                {
                    // Let the user know if this column is found in just the first
                    // query (1) or just the second query (2)
                    if (queryColumn.FirstQueryColName != null && queryColumn.SecondQueryColName != null)
                    {
                        if (queryColumn.CheckDataMatch)
                                columnAttribute = queryColumn.Tolerance == 0 ? "(M)" : string.Format("(M {0})", queryColumn.Tolerance);
                        else
                            columnAttribute = "";
                    }
                    else
                            columnAttribute = queryColumn.FirstQueryColName != null ? "(1)" : "(2)";

                    excelRow.Add(columnLetters[currColumnIndex].ToString(), new Cell(queryColumn.Label + " " + columnAttribute, CellStyle.DarkBlueBold));
                    ++currColumnIndex;                
                }
            }

            if (addRow)
            {
                rrExcelService.AddRow(excelRow);
            }
            return currColumnIndex;
        }

        /// <summary>
        /// Write the title and subtitle rows plus a blank row to the current worksheet. 
        /// Header rows will vary based on whether this is a one or two query recon.
        /// </summary>
        /// <param name="recon">The recon report we're processing</param>
        private void writeWorkSheetHeaderRows(ReconReport recon)
        {
            var firstQueryName = recon.FirstQueryName == recon.SecondQueryName ? recon.FirstQueryName + " (1)" : recon.FirstQueryName;
            var secondQueryName = recon.FirstQueryName == recon.SecondQueryName ? recon.SecondQueryName + " (2)" : recon.SecondQueryName;

            // Main title
            excelRow = new Dictionary<string, Cell>();
            excelRow.Add("A", new Cell(recon.Name, CellStyle.Bold));
            rrExcelService.AddRow(excelRow);
            // Sub-title row with names of queries
            excelRow.Clear();
            if (recon.SecondQueryName != "")
            {
                excelRow.Add("A", new Cell("Comparing"));
                excelRow.Add("B", new Cell(firstQueryName));
                excelRow.Add("C", new Cell("to"));
                excelRow.Add("D", new Cell(secondQueryName));
            }
            else
            {
                excelRow.Add("A", new Cell("Rows returned by"));
                excelRow.Add("B", new Cell(firstQueryName));
            }
            rrExcelService.AddRow(excelRow);
            // Write row with query variable values
            excelRow.Clear();
            excelRow.Add("A", new Cell("Query variables:"));
            var currColumnIndex = 1;
            if (recon.QueryVariables.Count > 0)
            {
                recon.QueryVariables.ForEach(queryVar =>
                    {
                        var queryVarName = queryVar.QuerySpecific ? string.Format("{0} ({1})", queryVar.SubName, queryVar.QueryNumber) : queryVar.SubName;
                        var currColumn = columnLetters[currColumnIndex].ToString();
                        excelRow.Add(currColumn, new Cell(string.Format("{0}: {1}", queryVarName, queryVar.SubValue)));
                        ++currColumnIndex;
                    });
            }
            else
                excelRow.Add("B", new Cell("None"));
            rrExcelService.AddRow(excelRow);

            // Sub-title row with key to column header names
            // Only add if two-query recon
            // The header text for matched column will vary depending on whether any non-zero tolerances exist
            // If none, just (M) = Matched Column, else (M x) = Matched column with tolerance of x
            if (secondQueryName != "")
            {
                string matchedColumnText = string.Empty;
                if (recon.Columns.FindAll(col => col.Tolerance > 0).Count() == 0)
                    matchedColumnText = "(M) = Matched Column";
                else
                    matchedColumnText = "(M x) = Matched Column and Tolerance";
                excelRow.Clear();
                excelRow.Add("A", new Cell("(Id) = Part of Unique ID", CellStyle.Italic));
                excelRow.Add("B", new Cell("(1) = 1st Query", CellStyle.Italic));
                excelRow.Add("C", new Cell("(2) = 2nd Query", CellStyle.Italic));
                excelRow.Add("D", new Cell(matchedColumnText, CellStyle.Italic));
                rrExcelService.AddRow(excelRow);
                /* ***** RESUME HERE: Add tolerance label to column headers for query 1/2 orphan sections ***** */
            }
            // Spacer row
            rrExcelService.AddBlankRow();
        }

        private void setReconQueryLabels(ReconReport recon)
        {
            if (recon.FirstQueryName != recon.SecondQueryName)
            {
                firstQueryLabel = recon.FirstQueryName;
                secondQueryLabel = recon.SecondQueryName;
            }
            else
            {
                firstQueryLabel = recon.FirstQueryName + " (1)";
                secondQueryLabel = recon.SecondQueryName + " (2)";
            }
        }

        #region Validation
        public bool ReadyToRun()
        {
            return (GetValidationErrors().Count == 0);
        }

        /// <summary>
        /// Perform various validation checks and return results for 
        /// any issues that are considered stoppers for running recons
        /// </summary>
        /// <returns>List with one line per error</returns>
        public List<string> GetValidationErrors()
        {
            var validationErrors = new List<string>();
            validationErrors.AddRange(validateSources());
            validationErrors.AddRange(validateRecons());  
            return validationErrors;
        }

        /// <summary>
        /// Identify any errors found with objects loaded from the Sources XML file. 
        /// Note that certain errors (missing entire sections) will halt any further error checking.
        /// </summary>
        /// <returns>List of error descriptions with one line per error</returns>
        private List<string> validateSources()
        {
            var sourcesMessages = new List<string>();
            var sourcesErrors = new List<string>();

            // First check if we have a possibly-complete set of info to check.  If not, don't go any further.
            if (sources == null || sources.ConnStringTemplates.Count == 0 || sources.ConnectionStrings.Count == 0 || sources.Queries.Count == 0)
            {
                sourcesErrors.Add("Sources: Incomplete information. Make sure at least one ConnectionStringTemplate, ConnectionString, and Query exists.");
                return sourcesErrors;
            }
            // Templates
            // No duplicate names
            var templateNames = (from csTemplate in sources.ConnStringTemplates
                                 select csTemplate.Name).ToList();
            var dupNames = getDuplicateEntries(templateNames);
            if (dupNames != string.Empty)
                sourcesErrors.Add(string.Format("Sources: Duplicate template name(s) found: {0}", dupNames));

            // Connection Strings
            // No duplicate names
            var connStringNames = (from connString in sources.ConnectionStrings
                                 select connString.Name).ToList();
            dupNames = getDuplicateEntries(connStringNames);
            if (dupNames != string.Empty)
                sourcesErrors.Add(string.Format("Sources: Duplicate connection string name(s) found: {0}", dupNames));
            sources.ConnectionStrings.ForEach(cs =>
            {
                // Template exists
                if (!templateNames.Contains(cs.TemplateName))
                    sourcesErrors.Add(string.Format("Sources: The connection string {0} refers to non-existent template {1}", cs.Name, cs.TemplateName));
                else
                {
                    // Template is valid, now make sure connection string supplies values for all placeholders and all placeholders are known
                    var templateString = sources.ConnStringTemplates.First(template => template.Name == cs.TemplateName).Template;
                    var matches = (from Match match in Regex.Matches(templateString, pipePlaceholderRegex)
                                   select match.Groups[1].Value).ToList();
                    var csVariableNames = (from csVariable in cs.TemplateVariables
                                           select csVariable.SubName).ToList();
                    matches.ForEach(placeholder =>
                    {
                        if (!csVariableNames.Contains(placeholder))
                            sourcesErrors.Add(string.Format("Sources: The connection string {0} is missing a value for the placeholder {1} in its template {2}", cs.Name, placeholder, cs.TemplateName));
                    });
                    csVariableNames.ForEach(variable =>
                    {
                        if (!matches.Contains(variable))
                            sourcesErrors.Add(string.Format("Sources: The connection string {0} has a variable {1} that does not have a placeholder in its template {2}", cs.Name, variable, cs.TemplateName));
                    });
                    // No duplicate connection string variable names
                    dupNames = getDuplicateEntries(csVariableNames);
                    if (dupNames != string.Empty)
                        sourcesErrors.Add(string.Format("Sources: The connection string {0} has duplicate variable names: {1}", cs.Name, dupNames));
                }
            });
            // Queries
            // Connection string exists
            sources.Queries.ForEach(query =>
            {
                if (!connStringNames.Contains(query.ConnStringName))
                    sourcesErrors.Add(string.Format("Sources: The query {0} refers to non-existent connection string {1}", query.Name, query.ConnStringName));
            });
            return sourcesErrors;
        }

        private List<string> validateRecons()
        {
            var reconsErrors = new List<string>();
            // First check if we have a possibly-complete set of info to check.  If not, don't go any further.
            if (recons == null || recons.ReconReports.Count() == 0)
            {
                reconsErrors.Add("Recons: Not loaded yet. A recon XML file w/at least one recon specified must be loaded.");
                return reconsErrors;
            }
            if (sources == null || sources.ConnectionStrings.Count() == 0)
            {
                reconsErrors.Add("Sources not loaded yet.");
                return reconsErrors;
            }
            var queryNames = (from query in sources.Queries select query.Name).ToList();
            // Duplicate recon names
            var dupNames = getDuplicateEntries((from recon in recons.ReconReports select recon.Name).ToList());
            if (dupNames != string.Empty)
                reconsErrors.Add(string.Format("Recons: Duplicate recon report name(s) found: {0}", dupNames));
            // Duplicate tab labels
            dupNames = getDuplicateEntries((from recon in recons.ReconReports select recon.TabLabel).ToList());
            if (dupNames != string.Empty)
                reconsErrors.Add(string.Format("Recons: Duplicate recon tab label(s) found: {0}", dupNames));
            List<string> allQueryPlaceholders = new List<string>();
            // Test for any query-specific issues for the recon
            recons.ReconReports.ForEach(recon =>
            {
                List<string> reconQueryNames = new List<string> { recon.FirstQueryName, recon.SecondQueryName };
                for (int queryNum = 1; queryNum <= 2; queryNum++)
                {
                    string reconQueryName = reconQueryNames[queryNum - 1];
                    // If query name is not empty, check query exists in source. If that's true continue on to check all related placeholders are given a value
                    if (reconQueryName != string.Empty)
                    {
                        if (!queryNames.Contains(reconQueryName))
                            reconsErrors.Add(string.Format("Recons: the recon {0} refers to a non-existent query called {1}", recon.Name, reconQueryName));
                        else
                        {
                            var querySql = sources.Queries.Find(query => query.Name == reconQueryName).SQL.Value;
                            var queryPlaceholders = (from Match match in Regex.Matches(querySql.ToString(), pipePlaceholderRegex)
                                                        select match.Groups[1].Value).ToList();
                            allQueryPlaceholders = allQueryPlaceholders.Union(queryPlaceholders).ToList();
                            var queryVariableNames = (from queryVariable in recon.QueryVariables
                                                        where !queryVariable.QuerySpecific || (queryVariable.QueryNumber == queryNum)
                                                        select queryVariable.SubName).ToList();
                            queryPlaceholders.ForEach(placeholder =>
                            {
                                if (!queryVariableNames.Contains(placeholder))
                                    reconsErrors.Add(string.Format("Recons: The recon {0} does not supply a placeholder value for {1} in query {2}", recon.Name, placeholder, reconQueryName));
                            });
                            // Make sure all variables listed as specific to the query have a corresponding placeholder in the related query
                            var querySpecificVarNames = (from queryVariable in recon.QueryVariables
                                                            where queryVariable.QueryNumber == queryNum
                                                            select queryVariable.SubName).ToList();
                            querySpecificVarNames.ForEach(queryVarName =>
                            {
                                if (!queryPlaceholders.Contains(queryVarName))
                                    reconsErrors.Add(string.Format("Recons: The recon {0} has a variable {1} for the query {2} with no corresponding placeholder in the query's sql", recon.Name, queryVarName, reconQueryName));
                            });
                            // No duplicate variable names for first query
                            dupNames = getDuplicateEntries(queryVariableNames);
                            if (dupNames != string.Empty)
                                reconsErrors.Add(string.Format("Recons: For recon {0} duplicate query variable(s) found {1} for first query {2}", recon.Name, dupNames, recon.FirstQueryName));
                        };
                    };
                };
                // Any non-query specific variables that are not used by either query
                var nonSpecificVarNames = (from queryVariable in recon.QueryVariables
                                            where !queryVariable.QuerySpecific
                                            select queryVariable.SubName).ToList();
                nonSpecificVarNames.ForEach(queryVarName =>
                {
                    if (!allQueryPlaceholders.Contains(queryVarName))
                        reconsErrors.Add(string.Format("Recons: The recon {0} has a non-query specific variable {1} with no corresponding placeholder any related SQL", recon.Name, queryVarName));
                });

                // Any column with non-zero tolerance must be a number and CheckDataMatch is true and tolerance must be a positive number
                var invalidColumns = (from col in recon.Columns
                                        where col.Tolerance < 0 || (col.Tolerance != 0 && (col.Type != ColumnType.number || !col.CheckDataMatch))
                                        select col.Label).ToList();
                if (invalidColumns.Count() > 0)
                {
                    var invalidColNames = string.Join(",", invalidColumns);
                    reconsErrors.Add(string.Format("Recons: The columns {0} are invalid because tolerances must be 0 for non-number columns, or >= 0 for number columns with CheckDataMatch = true. ", invalidColNames));
                }

                // Any variables that are QuerySpecific but don't have QueryNumber of 1 or 2
                var invalidQueryVariables = (from queryVariable in recon.QueryVariables
                                                where queryVariable.QuerySpecific && queryVariable.QueryNumber != 1 && queryVariable.QueryNumber != 2
                                                select queryVariable.SubName).ToList();
                if (invalidQueryVariables.Count() > 0)
                {
                    var invalidVarNames = string.Join(",", invalidQueryVariables);
                    reconsErrors.Add(string.Format("Recons: For recon {0} the query variable(s) {1} are invalid because marked query specific but have QueryNumber other than 1 or 2", recon.Name, invalidVarNames));
                }
                var identCols = (from col in recon.Columns
                                    where col.IdentifyingColumn == true
                                    select col.Label).ToList();
                var checkDataMatchCols = (from col in recon.Columns
                                            where col.CheckDataMatch == true
                                            select col.Label).ToList();
                if (recon.SecondQueryName == string.Empty)
                {
                    if (identCols.Count() > 0)
                    {
                        var identColList = string.Join(",", identCols);
                        reconsErrors.Add(string.Format("Recons: Recon {0} is a single-query recon but column(s) {1} are set as identifying column. ", recon.Name, identColList));
                    }
                    if (checkDataMatchCols.Count() > 0)
                    {
                        var dataMatchColList = string.Join(",", checkDataMatchCols);
                    reconsErrors.Add(string.Format("Recons: Recon {0} is a single-recon but column(s) {1} are set as needing to check if data matches between 2 queries.", recon.Name, dataMatchColList));
                    }                         
                }
                else
                {
                    // Tests specific to two-query recons
                    if (identCols.Count == 0)
                        reconsErrors.Add(string.Format("Recons: Recon {0} is a two-query recon but has no identifying columns to use to match rows between data sets.", recon.Name));
                    // Note: it's ok to have a two-query recon with no checkdatamatch columns because the  purpose of the recon may be only to identify orphans.
                }
            });
            return reconsErrors;
        }

        /// <summary>
        /// Identify duplicate entries in a list of strings
        /// </summary>
        /// <param name="stringList"></param>
        /// <returns>String with comma-delimited list of values appearing more than once, or string.empty if none</returns>
        private string getDuplicateEntries(List<string> stringList)
        {
            var dupStrings = stringList.GroupBy(str => str).ToDictionary(str => str.Key, str => str.Count()).Where(strCount => strCount.Value > 1).ToList();
            if (dupStrings.Count == 0)
                return string.Empty;
            else 
                return string.Join(",", dupStrings.Select(strCount => strCount.Key).ToList());
        }

        #endregion Validation

        #region Utilities
        /// <summary>
        /// If inner exception exists will append that message to top level exception message,
        /// else just returns top level exception message. Calls itself recursively in case
        /// there are >2 layers of exceptions.
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        public string GetFullErrorMessage(Exception ex)
        {
            if (ex.InnerException != null)
                return (string.Format("{0}: {1}", ex.Message, GetFullErrorMessage(ex.InnerException)));
            else
                return ex.Message;
        }

        /// <summary>
        /// Load sample recon and source data from the serializer rather than manually from files.
        /// Used for unit testing purposes.
        /// </summary>
        public void UseSampleData()
        {
            var rrFileService = getNewFileService();
            sources = rrFileService.SampleSources;
            recons = rrFileService.SampleRecons;
            rrDataService.Sources = sources;
            rrDataService.Recons = recons;
        }

        /// <summary>
        /// If we get an ActionStatusEvent from a service layer, just relay it along
        /// so the UI can act on it
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void relayActionStatusEvent(object sender, ActionStatusEventArgs e)
        {
            if (ActionStatusEvent != null)
                ActionStatusEvent(sender, e);
            if (e.State == RequestState.Error || e.State == RequestState.FatalError)
                throw new ApplicationException(string.Format("{0} halting on serious {1}: {2}", sender.ToString(), e.State, e.Message));
            else
                sendActionStatus(this, e.State, e.Message, false);
        }

        /// <summary>
        /// Use to send non-critical messages.  
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="state"></param>
        /// <param name="message"></param>
        /// <param name="isDetail">Set to false to be seen always, if true may only be seen if consumer of event has
        /// requested to see details or in detail logging.</param>
        private void sendActionStatus(object sender, RequestState state, string message, bool isDetail)
        {
            if (ActionStatusEvent != null)
                ActionStatusEvent(sender, new ActionStatusEventArgs(state, message, isDetail));
        }

        private RRFileService getNewFileService()
        {
            RRFileService rrFileService = new RRFileService();
            return rrFileService;
        }

        public ReconReport CopyReconReport(ReconReport reconReport)
        {
            return getNewFileService().CopyReconReport(reconReport);
        }

        #endregion Utilities
    }
}
