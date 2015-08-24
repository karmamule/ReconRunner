using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace ReconRunner.Model
{
    public class RRFileService
    {
        #region Sample builder
        private Recons sampleRecons;
        private Sources sampleSources;
        private Recons recons;
        private Sources sources;

        public RRFileService()
        {
            populateSampleReconObjects();
        }

        private void populateSampleReconObjects()
        {
            sampleSources = new Sources();
            sampleRecons = new Recons();
            // Create sample objects to use to populate sample Sources and RRReports XML files
            // for user to use as basis for real recons.

            // Create RRSourcesSample.xml
            // Contains sample connection template, connection, and query objects in SampleRRSources
            // Sample Connection String Templates
            var sqlServerTemplate = new ConnectionStringTemplate();
            sqlServerTemplate.Name = "SQL Server";
            sqlServerTemplate.Template = "Provider=SQLNCLI11;Server=|ServerName|;Database=|DatabaseName|;Trusted_Connection=yes";
            var teradataTemplate = new ConnectionStringTemplate();
            teradataTemplate.Name = "Teradata";
            teradataTemplate.Template = "Data Source=|ServerName|;User ID=|User|;Password=|Password|;Session Mode=TERADATA;";
            var oracleTemplate = new ConnectionStringTemplate();
            oracleTemplate.Name = "Oracle";
            oracleTemplate.Template = "Provider=MSDAORA.1;Data Source=|TNSName|;User Id=|User|;Password=|Password|;Persist Security Info=false;";
            sampleSources.ConnStringTemplates.Add(sqlServerTemplate);
            sampleSources.ConnStringTemplates.Add(teradataTemplate);
            sampleSources.ConnStringTemplates.Add(oracleTemplate);

            // Sample Connection Strings
            var warehouseConnectionString = new RRConnectionString();
            warehouseConnectionString.Name = "DataWarehouse Dev";
            warehouseConnectionString.TemplateName = "SQL Server";
            warehouseConnectionString.DatabaseType = DatabaseType.SQLServer;
            warehouseConnectionString.TemplateVariables = new List<Variable>
            {
                new Variable {SubName="ServerName", SubValue="bos-dbdevwh02"},
                new Variable {SubName="DatabaseName", SubValue="DataWarehouse" }
            };
            var dalUatConnectionString = new RRConnectionString();
            dalUatConnectionString.Name = "Teradata UAT";
            dalUatConnectionString.TemplateName = "Teradata";
            dalUatConnectionString.DatabaseType = DatabaseType.Teradata;
            dalUatConnectionString.TemplateVariables = new List<Variable>
            {
                new Variable {SubName="ServerName", SubValue="tddevbos"}, 
                new Variable {SubName="User", SubValue="DAL_READ"}, 
                new Variable {SubName="Password", SubValue="DAL_READ"}
            };

            var edmUatConnectionString = new RRConnectionString();
            edmUatConnectionString.Name = "EDM UAT";
            edmUatConnectionString.TemplateName = "Oracle";
            edmUatConnectionString.DatabaseType = DatabaseType.Oracle;
            edmUatConnectionString.TemplateVariables = new List<Variable>
            {
                new Variable {SubName="TNSName", SubValue="EDM_T"}, 
                new Variable {SubName="User", SubValue="edm_read"}, 
                new Variable {SubName="Password", SubValue="edm_read"}
            };
            sampleSources.ConnectionStrings.Add(dalUatConnectionString);
            sampleSources.ConnectionStrings.Add(edmUatConnectionString);
            sampleSources.ConnectionStrings.Add(warehouseConnectionString);

            // Sample Queries
            var doc = new XmlDocument();
            var portGicsSectorDetail = new RRQuery();
            portGicsSectorDetail.Name = "PortGicsSectorDetails";
            portGicsSectorDetail.ConnStringName = "DataWarehouse Dev";
            portGicsSectorDetail.SQL = doc.CreateCDataSection(@"with latestAttrInsEffDate as 
                                                            (
	                                                            select
		                                                            atins_id,
		                                                            max(effect_date) LatestRowDate
	                                                            from warehouse.attr_instrument
	                                                            where effect_date <= '|DateWanted|'
	                                                            group by atins_id
                                                            ),
                                                            latestAttrIns as
                                                            (
	                                                            select
		                                                            laied.atins_id,
		                                                            laied.LatestRowDate,
		                                                            max(ai.atins_release_number) LatestReleaseNumber
	                                                            from latestAttrInsEffDate laied
	                                                            join warehouse.attr_instrument ai
	                                                            on laied.atins_id = ai.atins_id
	                                                            and laied.LatestRowDate = ai.effect_date
	                                                            group by laied.atins_id, laied.LatestRowDate)
                                                            select 
	                                                            bigs.atptf_id,
	                                                            p.portf_name,
	                                                            bigs.calcul_date,
	                                                            bigs.atins_id,
	                                                            ai.atins_name,
	                                                            ai.effect_date,
	                                                            ai.end_valid_date,
                                                                ai.atins_release_number
                                                            from warehouse.bisamattributions_gics_sec bigs
                                                            join warehouse.portfolio p
                                                            on  bigs.atptf_id = |PortId| 
                                                            and bigs.atins_id <> 0
                                                            and bigs.calcul_date = '|DateWanted|'
                                                            and bigs.dateinactive = '31-dec-9999'
                                                            and bigs.atptf_id = p.portf_id
                                                            join latestAttrIns lai
                                                            on bigs.atins_id = lai.atins_id
															and bigs.atins_release_number = lai.LatestReleaseNumber
                                                            join warehouse.attr_instrument ai
                                                            on lai.atins_id = ai.atins_id
                                                            and lai.LatestRowDate = ai.effect_date
                                                            and lai.LatestReleaseNumber = ai.ATINS_RELEASE_NUMBER
                                                            order by atins_name");
            var portAttribGroupcheck = new RRQuery();
            portAttribGroupcheck.Name = "PortAttribGroupCheck";
            portAttribGroupcheck.ConnStringName = "DataWarehouse Dev";
            portAttribGroupcheck.SQL = doc.CreateCDataSection(@"with portAttributionsRuns as (
	                                                        select
		                                                        max(checksum(atptf_id, atins_id, atins_release_number, calcul_date, frequency_id, begin_date, portf_config_id)) portAttribRunKey,
		                                                        atptf_id,
		                                                        atins_id,
		                                                        atins_release_number,
		                                                        calcul_date, 
		                                                        frequency_id, 
		                                                        begin_date,
		                                                        portf_config_id,
		                                                        user_defined_group_id
	                                                        from warehouse.bisamAttributions_|SegmentTableName| bigs
	                                                        where atptf_id = |PortId| and atins_id <> 0 and dateinactive = '31-dec-9999'
	                                                        group by atptf_id, atins_id, atins_release_number, calcul_date, frequency_id, begin_date, portf_config_id, user_defined_group_id),
                                                        portMultiGroupCount as (
	                                                        select
		                                                        portAttribRunkey,
		                                                        atptf_id,
		                                                        atins_id,
		                                                        atins_release_number,
		                                                        calcul_date,
		                                                        frequency_id,
		                                                        begin_date,
		                                                        portf_config_id
	                                                        from portAttributionsRuns
	                                                        group by portAttribRunKey, atptf_id, atins_id, atins_release_number, calcul_date, frequency_id, begin_date, portf_config_id
	                                                        having count(*) > 1
                                                        )
                                                        select par.*, p.portf_name
                                                        from portAttributionsRuns par
                                                        join portMultiGroupCount pmgc
                                                        on par.portAttribRunKey = pmgc.portAttribRunKey
                                                        join warehouse.portfolio p
                                                        on par.atptf_id = p.portf_id
                                                        order by par.atins_id, par.atins_release_number, par.calcul_date, par.FREQUENCY_ID, par.begin_date, user_defined_group_id");
            var edmQuery = new RRQuery();
            edmQuery.Name = "EdmProductDayPositions";
            edmQuery.ConnStringName = "EDM UAT";
            edmQuery.SQL = doc.CreateCDataSection("select fm.snapshot_id, trim(fm.entity_id) entity_id, pd.security_alias, trim(pd.long_short_indicator) long_short_indicator, pd.market_value_base, pd.market_value_local from datamartdbo.fund_master fm join datamartdbo.position_details pd on fm.entity_id = '|ProductId|' and fm.effective_date = '|EdmDateWanted|' and fm.dmart_fund_id = pd.dmart_fund_id");
            var teraQuery = new RRQuery();
            teraQuery.Name = "DalProductDayPositions";
            teraQuery.ConnStringName = "Teradata UAT";
            teraQuery.SQL = doc.CreateCDataSection("select fp.snapshotid, dp.productid, dp.entitylongname, fp.OrigSecurityId, fp.LongShortIndicator, fp.marketvaluebase, fp.marketvaluelocal from dimproduct dp join factposition fp on dp.productid = '|ProductId|' and dp.dimproductid = fp.dimproductid and fp.dimtimeid = |DalDateWanted|");
            var posFkQuery = new RRQuery();
            posFkQuery.Name = "DalPositionsMissingFk";
            posFkQuery.ConnStringName = "Teradata UAT";
            posFkQuery.SQL = doc.CreateCDataSection("select * from (select dimtimeid,  case when dimsecurityid = 'UNKNOWN' and dimproductid = 'UNKNOWN' then 'BOTH' when dimsecurityid = 'UNKNOWN' then 'SECURITY' else 'PRODUCT' end MissingEntity, case when dimsecurityid = 'UNKNOWN' and dimproductid = 'UNKNOWN' then 'Security ID: ' || OrigSecurityId || '; Product ID: ' || OrigProductId when dimsecurityid = 'UNKNOWN' then OrigSecurityId else OrigProductId end MissingEntityId from factposition where (dimsecurityid = 'UNKNOWN' or dimproductid = 'UNKNOWN') and dimtimeid >= 1131101) data group by MissingEntity, MissingEntityId, DimTimeId order by MissingEntity, MissingEntityId, DimTimeId");
            sampleSources.Queries.Add(portGicsSectorDetail);
            sampleSources.Queries.Add(portAttribGroupcheck);
            sampleSources.Queries.Add(edmQuery);
            sampleSources.Queries.Add(teraQuery);
            sampleSources.Queries.Add(posFkQuery);

            // Create RRReportsSample.xml
            // Holds sample recon report definitions along with substitution values to use
            // in sample queries.
            // first recon report and write to file to create file to use as basis to create
            // full XML file specifying all the recon reports to run.

            // First two recons are examples of comparing two data sets
            // Acadian-specific example (Note: the actual example is not necessarily useful,
            // it's just being done to show how the process works)
            // This recon uses the same query twice, so the query variables are query specific
            ReconReport warehousePortGicsRecon = new ReconReport();
            warehousePortGicsRecon.Name = "Compare GICS Sec Details for 2 days for 1 portfolio";
            warehousePortGicsRecon.TabLabel = "GICS Sec Changes";
            warehousePortGicsRecon.FirstQueryName = "PortGicsSectorDetails";
            warehousePortGicsRecon.SecondQueryName = "PortGicsSectorDetails";

            QueryColumn portfolioId = new QueryColumn();
            portfolioId.Label = "Portfolio ID";
            portfolioId.Type = ColumnType.number;
            portfolioId.IdentifyingColumn = true;
            portfolioId.AlwaysDisplay = true;
            portfolioId.FirstQueryColName = "atptf_id";
            portfolioId.SecondQueryColName = "atptf_id";

            QueryColumn portfolioName = new QueryColumn();
            portfolioName.Label = "Portfolio Name";
            portfolioName.Type = ColumnType.text;
            portfolioName.IdentifyingColumn = false;
            portfolioName.AlwaysDisplay = true;
            portfolioName.FirstQueryColName = "portf_name";
            portfolioName.SecondQueryColName = "portf_name";

            QueryColumn calcDate = new QueryColumn();
            calcDate.Label = "Calcul Date";
            calcDate.Type = ColumnType.date;
            calcDate.IdentifyingColumn = false;
            calcDate.AlwaysDisplay = true;
            calcDate.FirstQueryColName = "calcul_date";
            calcDate.SecondQueryColName = "calcul_date";

            QueryColumn instrumentId = new QueryColumn();
            instrumentId.Label = "Instrument Id";
            instrumentId.Type = ColumnType.number;
            instrumentId.IdentifyingColumn = true;
            instrumentId.AlwaysDisplay = true;
            instrumentId.FirstQueryColName = "atins_id";
            instrumentId.SecondQueryColName = "atins_id";

            QueryColumn instrumentName = new QueryColumn();
            instrumentName.Label = "Instrument Name";
            instrumentName.Type = ColumnType.text;
            instrumentName.IdentifyingColumn = false;
            instrumentName.AlwaysDisplay = true;
            instrumentName.FirstQueryColName = "atins_name";
            instrumentName.SecondQueryColName = "atins_name";

            QueryColumn instrumentReleaseNum = new QueryColumn();
            instrumentReleaseNum.Label = "Release #";
            instrumentReleaseNum.Type = ColumnType.text;
            instrumentReleaseNum.IdentifyingColumn = false;
            instrumentReleaseNum.AlwaysDisplay = false;
            instrumentReleaseNum.CheckDataMatch = true;
            instrumentReleaseNum.FirstQueryColName = "atins_release_number";
            instrumentReleaseNum.SecondQueryColName = "atins_release_number";

            warehousePortGicsRecon.Columns.Add(portfolioId);
            warehousePortGicsRecon.Columns.Add(portfolioName);
            warehousePortGicsRecon.Columns.Add(calcDate);
            warehousePortGicsRecon.Columns.Add(instrumentId);
            warehousePortGicsRecon.Columns.Add(instrumentName);
            warehousePortGicsRecon.Columns.Add(instrumentReleaseNum);

            warehousePortGicsRecon.QueryVariables = new List<QueryVariable>
            {
                new QueryVariable { SubName = "PortId", SubValue = "1003022", QuerySpecific=false },
                new QueryVariable { SubName = "DateWanted", SubValue = "30-jul-2015", QuerySpecific=true, QueryNumber=1 },
                new QueryVariable { SubName = "DateWanted", SubValue = "31-jul-2015",QuerySpecific=true, QueryNumber=2 }
            };

            ReconReport instMultGroupsRecon = new ReconReport();
            instMultGroupsRecon.Name = "Instruments with Multiple Attribution Groups for Similar Runs";
            instMultGroupsRecon.TabLabel = "Multi-group Inst";
            instMultGroupsRecon.FirstQueryName = "PortAttribGroupCheck";
            instMultGroupsRecon.SecondQueryName = string.Empty;

            portfolioId = new QueryColumn();
            portfolioId.Label = "Portfolio ID";
            portfolioId.Type = ColumnType.number;
            portfolioId.AlwaysDisplay = true;
            portfolioId.FirstQueryColName = "atptf_id";

            portfolioName = new QueryColumn();
            portfolioName.Label = "Portfolio Name";
            portfolioName.Type = ColumnType.text;
            portfolioName.AlwaysDisplay = true;
            portfolioName.FirstQueryColName = "portf_name";

            instrumentId = new QueryColumn();
            instrumentId.Label = "Instrument Id";
            instrumentId.Type = ColumnType.number;
            instrumentId.AlwaysDisplay = true;
            instrumentId.FirstQueryColName = "atins_id";

            instrumentReleaseNum = new QueryColumn();
            instrumentReleaseNum.Label = "Release #";
            instrumentReleaseNum.Type = ColumnType.number;
            instrumentReleaseNum.AlwaysDisplay = true;
            instrumentReleaseNum.FirstQueryColName = "atins_release_number";

            calcDate = new QueryColumn();
            calcDate.Label = "Calcul Date";
            calcDate.Type = ColumnType.date;
            calcDate.AlwaysDisplay = true;
            calcDate.FirstQueryColName = "calcul_date";

            var frequencyId = new QueryColumn();
            frequencyId.Label = "Freq";
            frequencyId.Type = ColumnType.number;
            frequencyId.AlwaysDisplay = true;
            frequencyId.FirstQueryColName = "frequency_id";

            var beginDate = new QueryColumn();
            beginDate.Label = "Begin Date";
            beginDate.Type = ColumnType.date;
            beginDate.AlwaysDisplay = true;
            beginDate.FirstQueryColName = "begin_date";

            var portfConfigId = new QueryColumn();
            portfConfigId.Label = "Config Id";
            portfConfigId.Type = ColumnType.number;
            portfConfigId.AlwaysDisplay = true;
            portfConfigId.FirstQueryColName = "portf_config_id";

            var udGroupId = new QueryColumn();
            udGroupId.Label = "Group Id";
            udGroupId.Type = ColumnType.number;
            udGroupId.AlwaysDisplay = true;
            udGroupId.FirstQueryColName = "user_defined_group_id";

            instMultGroupsRecon.Columns.Add(portfolioId);
            instMultGroupsRecon.Columns.Add(portfolioName);
            instMultGroupsRecon.Columns.Add(instrumentId);
            instMultGroupsRecon.Columns.Add(instrumentReleaseNum);
            instMultGroupsRecon.Columns.Add(calcDate);
            instMultGroupsRecon.Columns.Add(frequencyId);
            instMultGroupsRecon.Columns.Add(beginDate);
            instMultGroupsRecon.Columns.Add(portfConfigId);
            instMultGroupsRecon.Columns.Add(udGroupId);

            instMultGroupsRecon.QueryVariables = new List<QueryVariable>
            {
                new QueryVariable { SubName = "PortId", SubValue = "1003022" },
                new QueryVariable { SubName = "SegmentTableName", SubValue = "GICS_Sec" }
            };

            sampleRecons.ReconReports.Add(warehousePortGicsRecon);
            sampleRecons.ReconReports.Add(instMultGroupsRecon);

            /*
            // Note that queries running on two separate platforms may still be compared
            ReconReport edmDalPortDayPositionRecon = new ReconReport();
            edmDalPortDayPositionRecon.Name = "EDM to DAL Product's Positions for Day";
            edmDalPortDayPositionRecon.TabLabel = "EDM DAL Prod Pos";
            edmDalPortDayPositionRecon.FirstQueryName = "EdmProductDayPositions";
            edmDalPortDayPositionRecon.SecondQueryName = "DalProductDayPositions";

            QueryColumn productId = new QueryColumn();
            productId.Label = "Product ID";
            productId.Type = ColumnType.text;
            productId.IdentifyingColumn = true;
            productId.AlwaysDisplay = true;
            productId.FirstQueryColName = "entity_id";
            productId.SecondQueryColName = "productid";

            QueryColumn fundName = new QueryColumn();
            fundName.Label = "Product Name";
            fundName.Type = ColumnType.text;
            fundName.AlwaysDisplay = true;
            fundName.FirstQueryColName = null;
            fundName.SecondQueryColName = "EntityLongName";

            QueryColumn snapshotId = new QueryColumn();
            snapshotId.Label = "Snapshot";
            snapshotId.Type = ColumnType.text;
            snapshotId.AlwaysDisplay = true;
            snapshotId.IdentifyingColumn = true;
            snapshotId.FirstQueryColName = "snapshot_id";
            snapshotId.SecondQueryColName = "snapshotid";

            QueryColumn securityId = new QueryColumn();
            securityId.Label = "Security ID";
            securityId.Type = ColumnType.number;
            securityId.IdentifyingColumn = true;
            securityId.AlwaysDisplay = true;
            securityId.FirstQueryColName = "security_alias";
            securityId.SecondQueryColName = "OrigSecurityId";

            QueryColumn longShortIndicator = new QueryColumn();
            longShortIndicator.Label = "Long-Short Indicator";
            longShortIndicator.Type = ColumnType.text;
            longShortIndicator.AlwaysDisplay = true;
            longShortIndicator.IdentifyingColumn = true;
            longShortIndicator.FirstQueryColName = "long_short_indicator";
            longShortIndicator.SecondQueryColName = "LongShortIndicator";

            QueryColumn marketValueBase = new QueryColumn();
            marketValueBase.Label = "Market Value Base";
            marketValueBase.Type = ColumnType.number;
            marketValueBase.CheckDataMatch = true;
            marketValueBase.FirstQueryColName = "market_value_base";
            marketValueBase.SecondQueryColName = "marketvaluebase";
            marketValueBase.Tolerance = 0;

            QueryColumn marketValueLocal = new QueryColumn();
            marketValueLocal.Label = "Market Value Local";
            marketValueLocal.Type = ColumnType.number;
            marketValueLocal.CheckDataMatch = true;
            marketValueLocal.FirstQueryColName = "market_value_local";
            marketValueLocal.SecondQueryColName = "marketvaluelocal";
            marketValueLocal.Tolerance = 0;

            edmDalPortDayPositionRecon.Columns.Add(productId);
            edmDalPortDayPositionRecon.Columns.Add(fundName);
            edmDalPortDayPositionRecon.Columns.Add(snapshotId);
            edmDalPortDayPositionRecon.Columns.Add(securityId);
            edmDalPortDayPositionRecon.Columns.Add(longShortIndicator);
            edmDalPortDayPositionRecon.Columns.Add(marketValueBase);
            edmDalPortDayPositionRecon.Columns.Add(marketValueLocal);

            edmDalPortDayPositionRecon.QueryVariables = new List<QueryVariable>
            { 
                new QueryVariable { SubName = "ProductId", SubValue = "EEUB", QuerySpecific=false }, 
                new QueryVariable { SubName = "EdmDateWanted", SubValue = "01-jul-2014", QuerySpecific=false }, 
                new QueryVariable { SubName = "DalDateWanted", SubValue = "1140701",QuerySpecific=false }                                                                          
            };
            sampleRecons.ReconReports.Add(edmDalPortDayPositionRecon);

            // This recon is an example of a recon with just one query, and any rows returned are assumed to
            // indicate an issue and will be reported
            ReconReport positionsMissingFkRecon = new ReconReport();
            positionsMissingFkRecon.Name = "Positions With Unknown Security or Product";
            positionsMissingFkRecon.TabLabel = "Pos Missing FK";
            positionsMissingFkRecon.FirstQueryName = "DalPositionsMissingFk";
            positionsMissingFkRecon.SecondQueryName = string.Empty;

            var dimTimeId = new QueryColumn();
            dimTimeId.Label = "Date";
            dimTimeId.Type = ColumnType.date;
            dimTimeId.FirstQueryColName = "dimtimeid";

            var missingEntity = new QueryColumn();
            missingEntity.Label = "Missing";
            missingEntity.Type = ColumnType.text;
            missingEntity.FirstQueryColName = "MissingEntity";

            var originalEntityId = new QueryColumn();
            originalEntityId.Label = "Entity Id(s)";
            originalEntityId.Type = ColumnType.text;
            originalEntityId.FirstQueryColName = "MissingEntityId";

            positionsMissingFkRecon.Columns.Add(dimTimeId);
            positionsMissingFkRecon.Columns.Add(missingEntity);
            positionsMissingFkRecon.Columns.Add(originalEntityId);

            sampleRecons.ReconReports.Add(positionsMissingFkRecon);
            */
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
            XmlSerializer serializer = new XmlSerializer(typeof(Sources));
            TextWriter writer = new StreamWriter(fileName);
            serializer.Serialize(writer, sampleSources);
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

        public void WriteSourcesToXMLFile(Sources sources, string fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Sources));
            TextWriter writer = new StreamWriter(fileName);
            serializer.Serialize(writer, sources);
            writer.Close();
        }

        /// <summary>
        /// Use to make a deep copy of a ReconReport
        /// </summary>
        /// <param name="reconReport">A ReconReport object to be copied</param>
        /// <returns>A deep copy of the ReconReport object provided</returns>
        public ReconReport CopyReconReport(ReconReport reconReport)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ReconReport));
                serializer.Serialize(stream, reconReport);
                stream.Position = 0;
                return (ReconReport)serializer.Deserialize(stream);
            }
        }

        public Recons ReadReconsFromXMLFile(string fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Recons));
            recons = new Recons();
            TextReader reader = new StreamReader(fileName);
            recons = (Recons)serializer.Deserialize(reader);
            return recons;
        }
        public Sources ReadSourcesFromXMLFile(string fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Sources));
            sources = new Sources();
            TextReader reader = new StreamReader(fileName);
            sources = (Sources)serializer.Deserialize(reader);
            return sources;
        }

        public Sources SampleSources
        {
            get { return sampleSources; }
        }

        public Recons SampleRecons
        {
            get { return sampleRecons; }
        }

    }
}
