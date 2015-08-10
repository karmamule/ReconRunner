using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace ReconRunner.Model
{
    public class RRSerializer
    {
        #region Sample builder
        private Recons sampleRecons = new Recons();
        private RRSources sampleRRSources = new RRSources();

        public RRSerializer()
        {
            populateSampleReconObjects();
        }

        private void populateSampleReconObjects()
        {
            // Create sample objects to use to populate sample RRSources and RRReports XML files
            // for user to use as basis for real recons.

            // Create RRSourcesSample.xml
            // Contains sample connection template, connection, and query objects in SampleRRSources

            // Sample Connection String Templates
            var teradataTemplate = new ConnectionStringTemplate();
            teradataTemplate.Name = "Teradata";
            teradataTemplate.Template = "Provider=TDOLEDB;Data Source=|ServerName|;Persist Security Info=True;User ID=|User|;Password=|Password|;Session Mode=TERA;";
            var oracleTemplate = new ConnectionStringTemplate();
            oracleTemplate.Name = "Oracle";
            oracleTemplate.Template = "Data Source=|TNSName|;User Id=|User|;Password=|Password|;Integrated Security=no;";
            var sampleTemplates = new ConnectionStringTemplates();
            sampleTemplates.TemplateList.Add(teradataTemplate);
            sampleTemplates.TemplateList.Add(oracleTemplate);

            // Sample Connection Strings
            var dalUatConnectionString = new RRConnectionString();
            dalUatConnectionString.Name = "Teradata UAT";
            dalUatConnectionString.Template = "Teradata";
            dalUatConnectionString.TemplateVariables = new List<Variable>{
                                                                                    new Variable {SubName="ServerName", SubValue="tddevbos"}, 
                                                                                    new Variable {SubName="User", SubValue="DAL_READ"}, 
                                                                                    new Variable {SubName="Password", SubValue="DAL_READ"}
                                                                              };

            var edmUatConnectionString = new RRConnectionString();
            edmUatConnectionString.Name = "EDM UAT";
            edmUatConnectionString.Template = "Oracle";
            edmUatConnectionString.TemplateVariables = new List<Variable>{
                                                                            new Variable {SubName="TNSName", SubValue="EDM_T"}, 
                                                                            new Variable {SubName="User", SubValue="edm_read"}, 
                                                                            new Variable {SubName="Password", SubValue="edm_read"}
                                                                         };
            var sampleConnStrings = new RRConnectionStrings();
            sampleConnStrings.RRConnStringList.Add(dalUatConnectionString);
            sampleConnStrings.RRConnStringList.Add(edmUatConnectionString);

            // Sample Queries
            var edmQuery = new RRQuery();
            edmQuery.Name = "EdmProductDayPositions";
            edmQuery.RRConnectionString = "EDM UAT";
            edmQuery.SQL = @"select fm.snapshot_id, fm.entity_id, pd.market_value_base, pd.market_value_local from datamartdbo.fund_master fm join datamartdbo.position_details pd on fm.entity_id = '|ProductId|' and fm.effective_date = '|EdmDateWanted|' and fm.dmart_fund_id = pd.dmart_fund_id;";
            var teraQuery = new RRQuery();
            teraQuery.Name = "DalProductDayPositions";
            teraQuery.RRConnectionString = "Teradata UAT";
            teraQuery.SQL = @"select fp.snapshotid, dp.productid, dp.entitylongname, fp.marketvaluebase, fp.marketvaluelocal from dimproduct dp join factposition fp on dp.productid = '|ProductId|' and dp.dimproductid = fp.dimproductid and fp.dimtimeid = |TeraDateWanted|;";
            var posFkQuery = new RRQuery();
            posFkQuery.Name = "DalPositionsMissingFk";
            posFkQuery.RRConnectionString = "Teradata UAT";
            posFkQuery.SQL = @"select * from (select dimtimeid,  case when dimsecurityid = 'UNKNOWN' and dimproductid = 'UNKNOWN' then 'BOTH' when dimsecurityid = 'UNKNOWN' then 'SECURITY' else 'PRODUCT' end MissingEntity, case when dimsecurityid = 'UNKNOWN' and dimproductid = 'UNKNOWN' then 'Security ID: '' || OrigSecurityId || ''; Product ID: '' || OrigProductId || '''  when dimsecurityid = 'UNKNOWN' then OrigSecurityId else OrigProductId end OriginalEntityId from factposition where (dimsecurityid = 'UNKNOWN' or dimproductid = 'UNKNOWN') and dimtimeid >= 1131101) data group by MissingEntity, OriginalEntityId, DimTimeId order by MissingEntity, OriginalEntityId, DimTimeId;";
            var sampleQueries = new Queries();
            sampleQueries.QueryList.Add(edmQuery);
            sampleQueries.QueryList.Add(teraQuery);
            sampleQueries.QueryList.Add(posFkQuery);

            sampleRRSources = new RRSources();
            sampleRRSources.ConnStringTemplates = sampleTemplates;
            sampleRRSources.ConnectionStrings = sampleConnStrings;
            sampleRRSources.Queries = sampleQueries;
            
            // Create RRReportsSample.xml
            // Holds sample recon report definitions along with substitution values to use
            // in sample queries.
            // first recon report and write to file to create file to use as basis to create
            // full XML file specifying all the recon reports to run.

            /* ****** RESUME TOMORROW: Add collection of VARIABLE values for queries.  
             *        Adapt sample recon report to use new RRSources definitions as above
             *        Test creation of sample XML files. */

            // First recon is an example of comparing two data sets from different databases
            ReconReport edmDalPortDayPositionRecon = new ReconReport();
            edmDalPortDayPositionRecon.Name = "EDM to DAL Product's Positions for Day";
            edmDalPortDayPositionRecon.TabLabel = "EDM DAL Pros Pos";
            edmDalPortDayPositionRecon.FirstQuery = "EdmProductDayPositions";
            edmDalPortDayPositionRecon.SecondQuery = "DalProductDayPositions";

            QueryColumn productId = new QueryColumn();
            productId.Label = "Product ID";
            productId.Type = ColumnType.text;
            productId.IdentifyingColumn = true;
            productId.AlwaysDisplay = true;
            productId.FirstQueryName = "entity_id";
            productId.SecondQueryName = "productid";

            QueryColumn fundName = new QueryColumn();
            fundName.Label = "Product Name";
            fundName.Type = ColumnType.text;
            fundName.AlwaysDisplay = true;
            fundName.FirstQueryName = "EntityLongName";
            fundName.SecondQueryName = null;

            QueryColumn snapshotId = new QueryColumn();
            snapshotId.Label = "Snapshot";
            snapshotId.Type = ColumnType.text;
            snapshotId.AlwaysDisplay = true;
            snapshotId.IdentifyingColumn = true;
            snapshotId.FirstQueryName = "snapshot_id";
            snapshotId.SecondQueryName = "snapshotid";

            QueryColumn marketValueBase = new QueryColumn();
            marketValueBase.Label = "Market Value Base";
            marketValueBase.Type = ColumnType.number;
            marketValueBase.ShouldMatch = true;
            marketValueBase.FirstQueryName = "market_value_base";
            marketValueBase.SecondQueryName = "marketvaluebase";

            QueryColumn marketValueLocal = new QueryColumn();
            marketValueLocal.Label = "Market Value Local";
            marketValueLocal.Type = ColumnType.number;
            marketValueLocal.ShouldMatch = true;
            marketValueLocal.FirstQueryName = "market_value_local";
            marketValueLocal.SecondQueryName = "marketvaluelocal";

            edmDalPortDayPositionRecon.Columns.Add(productId);
            edmDalPortDayPositionRecon.Columns.Add(fundName);
            edmDalPortDayPositionRecon.Columns.Add(snapshotId);
            edmDalPortDayPositionRecon.Columns.Add(marketValueBase);
            edmDalPortDayPositionRecon.Columns.Add(marketValueLocal);

            edmDalPortDayPositionRecon.QueryVariables = new List<Variable> { 
                                                                                new Variable { SubName = "ProductId", SubValue = "EEUB" }, 
                                                                                new Variable { SubName = "EdmDateWanted", SubValue = "01-jul-2014" }, 
                                                                                new Variable { SubName = "DalDateWanted", SubValue = "1140701" }                                                                          };
            sampleRecons.ReconList.Add(edmDalPortDayPositionRecon);

            // This second recon is an example of a recon with just one query, and any rows returned are assumed to
            // indicate an issue and will be reported
            ReconReport positionsMissingFkRecon = new ReconReport();
            positionsMissingFkRecon.Name = "Positions With Unknown Security or Product";
            positionsMissingFkRecon.TabLabel = "Pos Missing FK";
            positionsMissingFkRecon.FirstQuery = "DalPositionsMissingFk";
            positionsMissingFkRecon.SecondQuery = string.Empty;

            var dimTimeId = new QueryColumn();
            dimTimeId.Label = "Date";
            dimTimeId.Type = ColumnType.date;
            dimTimeId.AlwaysDisplay = true;
            dimTimeId.FirstQueryName = "dimtimeid";

            var date = new QueryColumn();
            date.Label = "Date";
            date.Type = ColumnType.date;
            date.AlwaysDisplay = true;
            date.FirstQueryName = "DimTimeId";

            var missingEntity = new QueryColumn();
            missingEntity.Label = "Missing";
            missingEntity.Type = ColumnType.text;
            missingEntity.AlwaysDisplay = true;
            missingEntity.FirstQueryName = "MissingEntity";

            var originalEntityId = new QueryColumn();
            originalEntityId.Label = "Entity Id";
            originalEntityId.Type = ColumnType.text;
            originalEntityId.AlwaysDisplay = true;
            originalEntityId.FirstQueryName = "OriginalEntityId";

            positionsMissingFkRecon.Columns.Add(dimTimeId);
            positionsMissingFkRecon.Columns.Add(missingEntity);
            positionsMissingFkRecon.Columns.Add(originalEntityId);

            // Don't add single query recon to sample file until supporting logic in controller added
            //sampleRecons.ReconList.Add(positionsMissingFkRecon);
        }

        public void WriteSampleReconsToXMLFile(string fileName)
        {
            // Write sample XML file holding sample recon reports and query variable substitution values
            XmlSerializer serializer = new XmlSerializer(typeof(Recons));
            TextWriter writer = new StreamWriter(fileName);
            serializer.Serialize(writer, sampleRecons);
            writer.Close();
        }

        public void WriteSampleSourcesToXMLFile(string fileName)
        {
            // Write sample XML file holding sample connection templates, connection strings, and queries
            XmlSerializer serializer = new XmlSerializer(typeof(RRSources));
            TextWriter writer = new StreamWriter(fileName);
            serializer.Serialize(writer, sampleRRSources);
            writer.Close();
        }
        #endregion Sample Builder

        public void WriteReconsToXMLFile(Recons recons, string fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Recons));
            TextWriter writer = new StreamWriter(fileName);
            serializer.Serialize(writer, recons);
            writer.Close();
        }

        public void WriteSourcesToXMLFile(RRSources sources, string fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(RRSources));
            TextWriter writer = new StreamWriter(fileName);
            serializer.Serialize(writer, sources);
            writer.Close();
        }

        public Recons ReadReconsFromXMLFile(string fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Recons));
            Recons recons = new Recons();
            TextReader reader = new StreamReader(fileName);
            return (Recons)serializer.Deserialize(reader);
        }
        public RRSources ReadSourcesFromXMLFile(string fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(RRSources));
            RRSources sources = new RRSources();
            TextReader reader = new StreamReader(fileName);
            return (RRSources)serializer.Deserialize(reader);
        }
    }
}
