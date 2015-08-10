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
        private Recons sampleRecons;
        private RRSources sampleRRSources;

        public RRSerializer()
        {
            populateSampleReconObjects();
        }

        private void populateSampleReconObjects()
        {
            sampleRRSources = new RRSources();
            sampleRecons = new Recons();
            // Create sample objects to use to populate sample RRSources and RRReports XML files
            // for user to use as basis for real recons.

            // Create RRSourcesSample.xml
            // Contains sample connection template, connection, and query objects in SampleRRSources
            // Sample Connection String Templates
            var teradataTemplate = new ConnectionStringTemplate();
            teradataTemplate.Name = "Teradata";
            teradataTemplate.Template = "Data Source=|ServerName|;User ID=|User|;Password=|Password|;Session Mode=TERADATA;";
            var oracleTemplate = new ConnectionStringTemplate();
            oracleTemplate.Name = "Oracle";
            oracleTemplate.Template = "Provider=MSDAORA.1;Data Source=|TNSName|;User Id=|User|;Password=|Password|;Persist Security Info=false;";
            sampleRRSources.ConnStringTemplates.Add(teradataTemplate);
            sampleRRSources.ConnStringTemplates.Add(oracleTemplate);

            // Sample Connection Strings
            var dalUatConnectionString = new RRConnectionString();
            dalUatConnectionString.Name = "Teradata UAT";
            dalUatConnectionString.TemplateName = "Teradata";
            dalUatConnectionString.DatabaseType = DatabaseType.Teradata;
            dalUatConnectionString.TemplateVariables = new List<Variable>{
                                                                                    new Variable {SubName="ServerName", SubValue="tddevbos"}, 
                                                                                    new Variable {SubName="User", SubValue="DAL_READ"}, 
                                                                                    new Variable {SubName="Password", SubValue="DAL_READ"}
                                                                              };

            var edmUatConnectionString = new RRConnectionString();
            edmUatConnectionString.Name = "EDM UAT";
            edmUatConnectionString.TemplateName = "Oracle";
            edmUatConnectionString.DatabaseType = DatabaseType.Oracle;
            edmUatConnectionString.TemplateVariables = new List<Variable>{
                                                                            new Variable {SubName="TNSName", SubValue="EDM_T"}, 
                                                                            new Variable {SubName="User", SubValue="edm_read"}, 
                                                                            new Variable {SubName="Password", SubValue="edm_read"}
                                                                         };
            sampleRRSources.ConnectionStrings.Add(dalUatConnectionString);
            sampleRRSources.ConnectionStrings.Add(edmUatConnectionString);

            // Sample Queries
            var edmQuery = new RRQuery();
            edmQuery.Name = "EdmProductDayPositions";
            edmQuery.ConnStringName = "EDM UAT";
            edmQuery.SQL = @"select fm.snapshot_id, trim(fm.entity_id) entity_id, pd.security_alias, trim(pd.long_short_indicator) long_short_indicator, pd.market_value_base, pd.market_value_local from datamartdbo.fund_master fm join datamartdbo.position_details pd on fm.entity_id = '|ProductId|' and fm.effective_date = '|EdmDateWanted|' and fm.dmart_fund_id = pd.dmart_fund_id";
            var teraQuery = new RRQuery();
            teraQuery.Name = "DalProductDayPositions";
            teraQuery.ConnStringName = "Teradata UAT";
            teraQuery.SQL = @"select fp.snapshotid, dp.productid, dp.entitylongname, fp.OrigSecurityId, fp.LongShortIndicator, fp.marketvaluebase, fp.marketvaluelocal from dimproduct dp join factposition fp on dp.productid = '|ProductId|' and dp.dimproductid = fp.dimproductid and fp.dimtimeid = |DalDateWanted|";
            var posFkQuery = new RRQuery();
            posFkQuery.Name = "DalPositionsMissingFk";
            posFkQuery.ConnStringName = "Teradata UAT";
            posFkQuery.SQL = @"select * from (select dimtimeid,  case when dimsecurityid = 'UNKNOWN' and dimproductid = 'UNKNOWN' then 'BOTH' when dimsecurityid = 'UNKNOWN' then 'SECURITY' else 'PRODUCT' end MissingEntity, case when dimsecurityid = 'UNKNOWN' and dimproductid = 'UNKNOWN' then 'Security ID: ' || OrigSecurityId || '; Product ID: ' || OrigProductId when dimsecurityid = 'UNKNOWN' then OrigSecurityId else OrigProductId end MissingEntityId from factposition where (dimsecurityid = 'UNKNOWN' or dimproductid = 'UNKNOWN') and dimtimeid >= 1131101) data group by MissingEntity, MissingEntityId, DimTimeId order by MissingEntity, MissingEntityId, DimTimeId";
            sampleRRSources.Queries.Add(edmQuery);
            sampleRRSources.Queries.Add(teraQuery);
            sampleRRSources.Queries.Add(posFkQuery);
            
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
            edmDalPortDayPositionRecon.TabLabel = "EDM DAL Prod Pos";
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
            fundName.FirstQueryName = null;
            fundName.SecondQueryName = "EntityLongName";

            QueryColumn snapshotId = new QueryColumn();
            snapshotId.Label = "Snapshot";
            snapshotId.Type = ColumnType.text;
            snapshotId.AlwaysDisplay = true;
            snapshotId.IdentifyingColumn = true;
            snapshotId.FirstQueryName = "snapshot_id";
            snapshotId.SecondQueryName = "snapshotid";

            QueryColumn securityId = new QueryColumn();
            securityId.Label = "Security ID";
            securityId.Type = ColumnType.number;
            securityId.IdentifyingColumn = true;
            securityId.AlwaysDisplay = true;
            securityId.FirstQueryName = "security_alias";
            securityId.SecondQueryName = "OrigSecurityId";

            QueryColumn longShortIndicator = new QueryColumn();
            longShortIndicator.Label = "Long-Short Indicator";
            longShortIndicator.Type = ColumnType.text;
            longShortIndicator.AlwaysDisplay = true;
            longShortIndicator.IdentifyingColumn = true;
            longShortIndicator.FirstQueryName = "long_short_indicator";
            longShortIndicator.SecondQueryName = "LongShortIndicator";

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
            edmDalPortDayPositionRecon.Columns.Add(securityId);
            edmDalPortDayPositionRecon.Columns.Add(longShortIndicator);
            edmDalPortDayPositionRecon.Columns.Add(marketValueBase);
            edmDalPortDayPositionRecon.Columns.Add(marketValueLocal);

            edmDalPortDayPositionRecon.QueryVariables = new List<Variable> { 
                                                                                new Variable { SubName = "ProductId", SubValue = "EEUB" }, 
                                                                                new Variable { SubName = "EdmDateWanted", SubValue = "01-jul-2014" }, 
                                                                                new Variable { SubName = "DalDateWanted", SubValue = "1140701" }                                                                          
            };
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
            dimTimeId.FirstQueryName = "dimtimeid";

            var missingEntity = new QueryColumn();
            missingEntity.Label = "Missing";
            missingEntity.Type = ColumnType.text;
            missingEntity.FirstQueryName = "MissingEntity";

            var originalEntityId = new QueryColumn();
            originalEntityId.Label = "Entity Id(s)";
            originalEntityId.Type = ColumnType.text;
            originalEntityId.FirstQueryName = "MissingEntityId";

            positionsMissingFkRecon.Columns.Add(dimTimeId);
            positionsMissingFkRecon.Columns.Add(missingEntity);
            positionsMissingFkRecon.Columns.Add(originalEntityId);

            sampleRecons.ReconList.Add(positionsMissingFkRecon);
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
