using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace ReconRunner.Model
{
    #region RRSources
    /// <summary>
    /// Essentially just a name-value pair to be used for substituting values for placeholders in either connection strings
    /// or queries.
    /// </summary>
    [XmlRoot("Variable")]
    public class Variable
    {
        string subName;
        string subValue;

        [XmlAttribute("SubName")]
        public string SubName
        {
            get { return subName; }
            set { subName = value; }
        }

        [XmlAttribute("SubValue")]
        public string SubValue
        {
            get { return subValue; }
            set { subValue = value; }
        }
    }

   #region RRQuery
    /// <summary>
    /// Query SQL can have pipe-delimited placeholders (similar to connection string templates) for variables that will be 
    /// substituted in at the time they're run.
    /// </summary>
    [XmlRoot("RRQuery")]
    public class RRQuery
    {
        string name;
        string rrConnStringName;
        string sql;

        [XmlAttribute ("Name")]
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        [XmlAttribute ("RRConnectionString")]
        public string RRConnectionString
        {
            get { return rrConnStringName; }
            set { rrConnStringName = value; }
        }

        /// <summary>
        /// Supply a SQL statement with one or more optional pipe-delimited variables to be specified at run time.
        /// </summary>
        [XmlAttribute ("SQL")]
        public string SQL
        {
            get { return sql; }
            set { sql = value; }
        }
    }

    /// <summary>
    /// A collection of ReconReport objects to be run.  Each of them will have a tab created showing
    /// their results in the final spreadsheet created.
    /// </summary>
    [XmlRoot("Queries")]
    public class Queries
    {
        List<RRQuery> queryList;

        public Queries()
        {
            queryList = new List<RRQuery>();
        }

        /// <summary>
        /// A Collection of recon reports to be run.  Note that their key integer values will be used to control
        /// the order they're run in.
        /// </summary>
        [XmlElement("Query")]
        public List<RRQuery> QueryList
        {
            get { return queryList; }
            set { queryList = value; }
        }

    }
    #endregion RRQuery

    #region ConnectionStringTemplate and collection
    /// <summary>
    /// The general intended usage is one ConnectionStringTemplate for database type (Oracle, Teradata, etc.)  However if you want
    /// to have something vary that is not currently coded to be variable for each connectionstring instance then you'll either need to 
    /// make a new template or adjust what is variable by template in the code here.  
    /// 
    /// Put Pipe character "|" around any named variable that will be filled by a Variable.  Any connection string using this template must provide a corresponding
    /// value in its TemplateVariables collection.
    /// </summary>
    [XmlRoot("Template")]
    public class ConnectionStringTemplate
    {
        string name;
        string template;

        [XmlAttribute ("Template")]
        public string Template
        {
            get { return template; }
            set { template = value; }
        }

        [XmlAttribute ("Name")]
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
    }

    [XmlRoot("ConnectionStringTemplates")]
    public class ConnectionStringTemplates
    {
        List<ConnectionStringTemplate> templateList;

        public ConnectionStringTemplates()
        {
            templateList = new List<ConnectionStringTemplate>();
        }
        

        [XmlElement ("ConnectionStringTemplate")]
        public List<ConnectionStringTemplate> TemplateList
        {
            get { return templateList; }
            set { templateList = value; }
        }
    }         
    #endregion ConnectionStringTemplate and collection

    #region RRConnectionString and collection
    /// <summary>
    /// Each connection string will be an instance of the template with corresponding variable values to plug in
    /// for the appropriate placeholders as delimited by | in the template
    /// </summary>
    [XmlRoot("Template")]
    public class RRConnectionString
    {
        string name;
        string template;
        List<Variable> templateVariables;

        [XmlAttribute("Name")]
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        [XmlAttribute("Template")]
        public string Template
        {
            get { return template; }
            set { template = value; }
        }

        [XmlElement("TemplateVariable")]
        public List<Variable> TemplateVariables
        {
            get { return templateVariables; }
            set { templateVariables = value; }
        }
    }

    [XmlRoot("RRConnectionStrings")]
    public class RRConnectionStrings
    {
        List<RRConnectionString> rrConnStringList;

        public RRConnectionStrings()
        {
            rrConnStringList = new List<RRConnectionString>();
        }


        [XmlElement("RRConnection")]
        public List<RRConnectionString> RRConnStringList
        {
            get { return rrConnStringList; }
            set { rrConnStringList = value; }
        }
    }
    #endregion RRConnectionString and collection

    [XmlRoot("RRSources")]
    public class RRSources
    {
        private ConnectionStringTemplates csTemplates;
        private RRConnectionStrings connStrings;
        private Queries queries;

        [XmlElement("ConnectionStringTemplate")]
        public ConnectionStringTemplates ConnStringTemplates
        {
            get { return csTemplates; }
            set { csTemplates = value; }
        }

        [XmlElement("ConnectionString")]
        public RRConnectionStrings ConnectionStrings
        {
            get { return connStrings; }
            set { connStrings = value; }
        }

        [XmlElement("Queries")]
        public Queries Queries
        {
            get { return queries; }
            set { queries = value; }
        }
    }

    #endregion RRSources

#region RRReports

    #region ColumnType
    public enum ColumnType
    {
        text,
        date,
        number
    }
    #endregion ColumnType

    #region QueryColumn
    [XmlRoot("QueryColumn")]
    public class QueryColumn
    {
        string label;
        string firstQueryName;
        string secondQueryName;
        ColumnType type;
        bool shouldMatch;
        bool alwaysDisplay;
        bool identifyingColumn;

        public QueryColumn()
        {
            type = ColumnType.text;
            shouldMatch = false;
            alwaysDisplay = false;
            identifyingColumn = false;
        }

        /// <summary>
        /// The name for this column.  Will be used as column title in the resulting
        /// recon spreadsheet.
        /// </summary>
        [XmlAttribute("Label")]
        public string Label
        {
            get { return label; }
            set { label = value; }
        }

        /// <summary>
        /// What this QueryColumn is called in the first query.  Set to NULL if the first query
        /// doesn't have this column.  (Note that shouldMatch must be FALSE if set to NULL.)
        /// Query name will be used as part of subtitle in recon report created.
        /// </summary>
        [XmlAttribute("FirstQueryName")]
        public string FirstQueryName
        {
            get {return firstQueryName; }
            set { firstQueryName = value; }
        }

        /// <summary>
        /// What this QueryColumn is called in the second query.  Set to NULL if the second query
        /// doesn't have this column.  (Note that shouldMatch must be FALSE if set to NULL.)
        /// Query name will be used as part of subtitle in recon report created.
        /// </summary>
        [XmlAttribute("SecondQueryName")]
        public string SecondQueryName
        {
            get { return secondQueryName; }
            set { secondQueryName = value; }
        }

        /// <summary>
        /// The ColumnType associated with this column.  Uses an enumerator
        /// ColumnType with acceptable values of text, date, or number.
        /// Defaults to text.
        /// </summary>
        [XmlAttribute("Type")]
        public ColumnType Type
        {
            get { return type; }
            set { type = value; }
        }

        /// <summary>
        /// A boolean indicating if the column should be compared to its equivalent
        /// in both queries being compared.   Defaults to false.
        /// </summary>
        [XmlAttribute("ShouldMatch")]
        public bool ShouldMatch
        {
            get { return shouldMatch; }
            set {shouldMatch = value;}
        }

        /// <summary>
        /// A boolean indicating the column should appear on any recon row created
        /// regardless of whether or not it was the column with a recon issue.  Defaults to false;
        /// </summary>
        [XmlAttribute("AlwaysDisplay")]
        public bool AlwaysDisplay
        {
            get { return alwaysDisplay; }
            set { alwaysDisplay = value; }
        }

        /// <summary>
        /// Set to true if this is one of the columns that will be used to identify what records
        /// should be compared.  Two records will only be reconned against each other if all of their
        /// identifying columns are the same.  Defaults to false;
        /// </summary>
        [XmlAttribute("IdentifyingColumn")]
        public bool IdentifyingColumn
        {
            get { return identifyingColumn; }
            set { identifyingColumn = value; }
        }
    }
    #endregion QueryColumn

    /// <summary>
    /// A ReconReport has a set of columns and two queries.  The assumption is the two queries have most columns in common.
    /// Each query may optionally have a list of string substitution values
    /// </summary>
    [XmlRoot("ReconReport")]
    public class ReconReport
    {
        string name;
        string tabLabel;
        List<QueryColumn> columns;
        string firstQuery;
        List<Variable> queryVariables;
        string secondQuery;

        public ReconReport()
        {
            columns = new List<QueryColumn>();
        }

        /// <summary>
        /// What to title this recon report.  Will be first line of tab created for this report
        /// </summary>
        [XmlAttribute("Name")]
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        /// <summary>
        /// What to use for the label of the Tab for the recon report that will be created
        /// in the resulting spreadsheet.
        /// </summary>
        [XmlAttribute("TabLabel")]
        public string TabLabel
        {
            get { return tabLabel; }
            set { tabLabel = value; }
        }

        /// <summary>
        /// A list of QueryColumn objects that each query in the recon will return.
        /// </summary>
        [XmlElement("QueryColumn")]
        public List<QueryColumn> Columns
        {
            get { return columns; }
            set { columns = value; }
        }

        [XmlElement("QueryVariable")]
        public List<Variable> QueryVariables
        {
            get { return queryVariables; }
            set { queryVariables = value; }
        }        
        
        /// <summary>
        /// A ReconQuery object for the first query in this recon
        /// </summary>
        [XmlAttribute("FirstQuery")]
        public string FirstQuery
        {
            get { return firstQuery; }
            set { firstQuery = value; }
        }

        /// <summary>
        /// A ReconQuery object for the second query in this recon.
        /// </summary>
        [XmlAttribute("SecondQuery")]
        public string SecondQuery
        {
            get { return secondQuery; }
            set { secondQuery = value; }
        }


    }

    /// <summary>
    /// A collection of ReconReport objects to be run.  Each of them will have a tab created showing
    /// their results in the final spreadsheet created.
    /// </summary>
    [XmlRoot("Recons")]
    public class Recons
    {
        List<ReconReport> reconList;

        public Recons() 
        {
            reconList = new List<ReconReport>();
        }

        /// <summary>
        /// A Collection of recon reports to be run.  Note that their key integer values will be used to control
        /// the order they're run in.
        /// </summary>
        [XmlElement("ReconReport")]
        public List<ReconReport> ReconList
        {
            get { return reconList; }
            set { reconList = value; }
        }

    }

#endregion RRReports
}
