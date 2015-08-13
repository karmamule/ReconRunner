using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using System.Data.Common;
using System.Data.OleDb;
using Teradata.Client.Provider;

namespace ReconRunner.Model
{
    public class RRDataService
    {
        #region Constructors
        private static RRDataService instance = new RRDataService();
        public static RRDataService Instance { get { return instance; } }
        private static RRSerializer serializer = new RRSerializer();

        private Sources sources;
        private Recons recons;
        // A set of open connections available for use
        // Use OpenConnections() and CloseConnections() to manage them for a given recon
        // Note connections can be of type TdConnection (for Teradata) or OleDbConnection (for anything else)
        private Dictionary<string, DbConnection> openConnections = new Dictionary<string, DbConnection>();

        private RRDataService()
        {

        }

        
        /// <summary>
        /// The method OpenConnections() should be called before attempting any
        /// call to GetReconData.  A list of datatables will be returned, with the
        /// first entry being the data for the first query, and if there is a second query
        /// then the second entry will be its datatable.
        /// 
        /// When done CloseConnections() should be called.
        /// </summary>
        /// <param name="recon"></param>
        /// <returns></returns>
        public List<DataTable> GetReconData(ReconReport recon)
        {
            if (sources == null)
                throw new Exception("RRDataService: Sources not provided yet");
            var reconData = new List<DataTable>();
            try
            {
                // Get data for first query
                var firstQuery = sources.Queries.Single(query => query.Name == recon.FirstQueryName);
                reconData.Add(getQueryDataTable(firstQuery, recon.QueryVariables));
                // If necessary, get data for second query
                if (recon.SecondQueryName != "")
                {
                    var secondQuery = sources.Queries.Single(query => query.Name == recon.SecondQueryName);
                    reconData.Add(getQueryDataTable(secondQuery, recon.QueryVariables));
                }
                return reconData;
            }
            catch (Exception ex)
            {
                CloseConnections();
                throw new Exception("Problem running first query: " + getFullErrorMessage(ex));
            }
        }

        /// <summary>
        /// Get the datatable for the provided query and substitution variables.  Note that OpenConnections()
        /// should have been run before this is run.
        /// </summary>
        /// <param name="query">The RRQuery object to be executed</param>
        /// <param name="queryVariables">Substitution variables to be used to create the final SQL to be run</param>
        /// <returns></returns>
        private DataTable getQueryDataTable(RRQuery query, List<QueryVariable> queryVariables)
        {
            var dataSet = new DataSet();
            // Get the connection for the query.     
            var firstQueryDbConn = openConnections.Single(conn => conn.Key == query.ConnStringName).Value;
            // Get the query's sql and perf
            var firstQuerySql = buildQuerySql(queryVariables, query.SQL.Value);
            var connType = firstQueryDbConn.GetType();
            if (connType.Name == "OleDbConnection")
            {
                // It's an Ole DB connection
                var firstQueryConn = (OleDbConnection)firstQueryDbConn;
                using (OleDbCommand command = new OleDbCommand(firstQuerySql, firstQueryConn))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    adapter.Fill(dataSet);
                return dataSet.Tables[0];
            }
            else
            {
                // It's a teradata connection
                var firstQueryConn = (TdConnection)firstQueryDbConn;
                using (TdCommand command = new TdCommand(firstQuerySql, firstQueryConn))
                using (TdDataAdapter adapter = new TdDataAdapter(command))
                    adapter.Fill(dataSet);
                return dataSet.Tables[0];
            }
        }

        /// <summary>
        /// Go through all the queries used in recons and open their corresponding connections
        /// </summary>
        /// <param name="recons">A list of recons to be done</param>
        public void OpenConnections(Recons recons)
        {
            try
            {
                recons.ReconReports.ForEach(openReconConnections);
            }   
            catch (Exception ex)
            {
                CloseConnections();
                throw ex;
            }

        }

        /// <summary>
        /// Look at one or two connections required by the recon report.  If they're already
        /// present in openConnections then do nothing, otherwise create a connection and add it 
        /// to the collection.
        /// </summary>
        /// <param name="rr">A recon report</param>
        private void openReconConnections(ReconReport rr)
        {
                // Open connection for first query if needed
                try
                {
                    var firstConnection = sources.ConnectionStrings.Where(cs => cs.Name == sources.Queries.Where(query => query.Name == rr.FirstQueryName).First().ConnStringName).First();
                    openConnection(firstConnection);
                    if (rr.SecondQueryName != "")
                    {
                        var secondConnection = sources.ConnectionStrings.Where(cs => cs.Name == sources.Queries.Where(query => query.Name == rr.SecondQueryName).First().ConnStringName).First();
                        openConnection(secondConnection);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Error while trying to open connection for report {0}: {1}", rr.Name, getFullErrorMessage(ex)));
                }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="connectionName">Name of connection (should be held in sources connections collection) to be opened
        /// and placed in openConnections if not already there.  (If it is already there then nothing is done)</param>
        /// <returns></returns>
        private void openConnection(RRConnectionString connection)
        {
            if (openConnections.ContainsKey(connection.Name))
                // Have already opened that connection, do nothing
                return;
            else
            {
                // Have not yet opened that connection, so open it and save to collection
                var connString = buildConnectionString(connection.Name);
                if (connection.DatabaseType == DatabaseType.Teradata)
                {
                    var openConnection = new TdConnection(connString);
                    openConnection.Open();
                    openConnections.Add(connection.Name, openConnection);
                }
                else
                {
                    var openConnection = new OleDbConnection(connString);
                    openConnection.Open();
                    var connType = openConnection.GetType();
                    openConnections.Add(connection.Name, openConnection);
                }
            }
        }

        public void CloseConnections()
        {
            try
            {
                openConnections.Values.ToList().ForEach(conn =>
                {
                    conn.Close();
                    conn.Dispose();
                });
                openConnections.Clear();
            } 
            catch (Exception ex)
            {
                throw new Exception("Error while closing connections: " + getFullErrorMessage(ex));
            }
        }

        /// <summary>
        /// Take name of conn string, get it from sources, and mix substitution values of that connection string
        /// with its corresponding template to generate the actual connection string
        /// </summary>
        /// <param name="connStringName">Name of a connection string in sources to be built from template</param>
        private string buildConnectionString(string connStringName)
        {
            var connString = sources.ConnectionStrings.Where(cs => cs.Name == connStringName).First();
            var connStringTemplate = sources.ConnStringTemplates.Where(t => t.Name == connString.TemplateName).First();
            var finalConnString = connStringTemplate.Template;
            connString.TemplateVariables.ForEach(variable =>
            {
                finalConnString = finalConnString.Replace("|" + variable.SubName + "|", variable.SubValue);
            });
            return finalConnString;
        }

        /// <summary>
        /// Takes base SQL for a query, a set of variables, and performs substitutions to
        /// create SQL that will actually be run
        /// </summary>
        /// <param name="reconVariables"></param>
        /// <param name="querySql"></param>
        /// <returns></returns>
        private string buildQuerySql(List<QueryVariable> reconVariables, string querySql)
        {
            string finalQuerySql = querySql;
            reconVariables.ForEach(variable =>
                {
                    finalQuerySql = finalQuerySql.Replace("|" + variable.SubName + "|", variable.SubValue);
                });
            return finalQuerySql;
        }

        #endregion Constructors


        #region utilities
        public DataTable ConvertReaderToTable(DbDataReader reader)
        {
            DataReaderAdapter dra = new DataReaderAdapter();
            DataTable dt = new DataTable();
            dra.FillFromReader(dt, reader);
            return dt;
        }

        /// <summary>
        /// If inner exception exists will append that message to top level exception message,
        /// else just returns top level exception message. Calls itself recursively in case
        /// there are >2 layers of exceptions.
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        private string getFullErrorMessage(Exception ex)
        {
            if (ex.InnerException != null)
                return (string.Format("{0}: {1}", ex.Message, getFullErrorMessage(ex.InnerException)));
            else
                return ex.Message;
        }

        #endregion utilities
    }
}
