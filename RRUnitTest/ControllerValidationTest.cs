using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReconRunner.Model;
using ReconRunner.Controller;


namespace RRUnitTest
{
    [TestClass]
    public class ControllerValidationTest
    {
        RRController rrController = RRController.Instance;

        #region Test Validation Logic
        [TestMethod]
        public void SampleDataValid()
        {
            rrController.UseSampleData();
            Assert.IsTrue(rrController.ReadyToRun(), "Sample data invalid: " + string.Join(",", rrController.GetValidationErrors()));           
        }

        [TestMethod]
        public void DetectNoConnections()
        {
            rrController.UseSampleData();
            rrController.Sources.ConnectionStrings.Clear();
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect when no connection strings.");
        }

        [TestMethod]
        public void DetectDupTemplateNames()
        {
            rrController.UseSampleData();
            rrController.Sources.ConnStringTemplates[0].Name = rrController.Sources.ConnStringTemplates[1].Name;
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect duplicate connection string template names.");
        }

        [TestMethod]
        public void DetectUnknownTemplate()
        {
            rrController.UseSampleData();
            rrController.Sources.ConnectionStrings[0].TemplateName += "XYZ";
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a connection string using an unknown template name.");
        }

        [TestMethod]
        public void DetectConnStringMissingPlaceholder()
        {
            rrController.UseSampleData();
            rrController.Sources.ConnectionStrings[0].TemplateVariables.RemoveAt(0);
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a connection string was missing a required placeholder value in the template.");
        }

        [TestMethod]
        public void DetectDupConnStringName()
        {
            rrController.UseSampleData();
            rrController.Sources.ConnectionStrings[0].Name = rrController.Sources.ConnectionStrings[1].Name;
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect duplicate connection string names.");
        }

        [TestMethod]
        public void DetectUnknownConnStringPlaceholder()
        {
            rrController.UseSampleData();
            rrController.Sources.ConnectionStrings[0].TemplateVariables.Add(new Variable { SubName = "XYZ123987", SubValue = "0" });
            var errors = rrController.GetValidationErrors();
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect conn string with placeholder name not referenced in template.");
        }

        [TestMethod]
        public void DetectDupTemplateVariables()
        {
            rrController.UseSampleData();
            rrController.Sources.ConnectionStrings[0].TemplateVariables.Add(rrController.Sources.ConnectionStrings[0].TemplateVariables[0]);
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect multiple instances of same template variable for connection string.");
        }

        [TestMethod]
        public void DetectUnknownConnString()
        {
            rrController.UseSampleData();
            rrController.Sources.Queries[0].ConnStringName += "XYZ";
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect query using an unknown connection string name.");
        }

        [TestMethod]
        public void DetectDupReconNames()
        {
            rrController.UseSampleData();
            rrController.Recons.ReconReports[0].Name = rrController.Recons.ReconReports[1].Name;
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect duplicate recon report names.");
        }

        [TestMethod]
        public void DetectDupTabLabels()
        {
            rrController.UseSampleData();
            rrController.Recons.ReconReports[0].TabLabel = rrController.Recons.ReconReports[1].TabLabel;
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect duplicate recon tab labels.");
        }

        [TestMethod]
        public void DetectUnknownQuery()
        {
            rrController.UseSampleData();
            rrController.Recons.ReconReports[0].FirstQueryName += "XYZ";
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect unknown name for first query.");
            rrController.UseSampleData();
            rrController.Recons.ReconReports.Find(recon => recon.SecondQueryName != string.Empty).SecondQueryName += "XYZ";
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect unknown name for second query.");
        }

        [TestMethod]
        public void DetectReconMissingPlaceholder()
        {
            rrController.UseSampleData();
            var testRecon = rrController.Recons.ReconReports.Find(recon => recon.QueryVariables.Count > 1);
            rrController.Recons.ReconReports.Find(recon => recon.QueryVariables.Count > 1).QueryVariables.RemoveAt(0);
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect recon missing a variable for a query placeholder");
        }

        [TestMethod]
        public void DuplicateSubNames()
        {
            rrController.UseSampleData();
            var testRecon = rrController.Recons.ReconReports.Find(recon => recon.QueryVariables.Count > 1);
            testRecon.QueryVariables.Add(testRecon.QueryVariables[0]);
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect duplicate placeholder variable for a recon.");
        }

        [TestMethod]
        public void DetectInvalidQueryVariable()
        {
            rrController.UseSampleData();
            var testRecon = rrController.Recons.ReconReports.Find(recon => recon.QueryVariables.Exists(qv => qv.QuerySpecific == true));
            var testQv = testRecon.QueryVariables.Find(qv => qv.QuerySpecific == true);
            testQv.QueryNumber = 0;
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a QuerySpecific query variable whose querynumber was not 1 or 2.");
        }

        [TestMethod]
        public void DetectProperColumnTypes()
        {
            // Identifying columns are not used by single-query recons, and at least one needed for two-query recons.
            rrController.UseSampleData();
            rrController.Recons.ReconReports.Find(recon => recon.SecondQueryName == string.Empty).Columns[0].IdentifyingColumn = true;
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a single-query recon that has an identifying column.");
            rrController.UseSampleData();
            var testRecon = rrController.Recons.ReconReports.Find(recon => recon.SecondQueryName != string.Empty);
            testRecon.Columns.ForEach(col => col.IdentifyingColumn = false);
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a two-query recon with no identifying columns.");
            // CheckDataMatch columns are not needed for single query recons. (And optional for two-query recons)
            rrController.UseSampleData();
            rrController.Recons.ReconReports.Find(recon => recon.SecondQueryName == string.Empty).Columns[0].CheckDataMatch = true;
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a single-query recon that has a CheckDataMatch column.");
        }

        [TestMethod]
        public void DetectUnknownQueryPlaceholder()
        {
            rrController.UseSampleData();
            rrController.Recons.ReconReports[0].QueryVariables.Add(new QueryVariable { SubName = "XYZ123987", SubValue = "0", QuerySpecific = false, QueryNumber = 0 });
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a non-query specific recon placeholder not referenced by any query.");
            rrController.UseSampleData();
            rrController.Recons.ReconReports[0].QueryVariables.Add(new QueryVariable { SubName = "XYZ123987", SubValue = "0", QuerySpecific = true, QueryNumber = 1 });
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a first-query specific recon placeholder not referenced in specified query.");
            rrController.UseSampleData();
            rrController.Recons.ReconReports.Find(recon => recon.SecondQueryName != string.Empty).QueryVariables.Add(new QueryVariable { SubName = "XYZ123987", SubValue = "0", QuerySpecific = true, QueryNumber = 2 });
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a second-query specific recon placeholder not referenced in specified query.");
        }

        [TestMethod]
        public void DetectInvalidTolerance()
        {
            rrController.UseSampleData();
            var testRecon = rrController.Recons.ReconReports.Find(recon => recon.Columns.Exists(col => col.Type != ColumnType.number));
            var testQc = testRecon.Columns.Find(col => col.Type != ColumnType.number);
            testQc.Tolerance = 1;
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a non-number column having a non-zero tolerance.");
            rrController.UseSampleData();
            testRecon = rrController.Recons.ReconReports.Find(recon => recon.Columns.Exists(col => col.Type == ColumnType.number && !col.CheckDataMatch));
            testQc = testRecon.Columns.Find(col => col.Type == ColumnType.number && !col.CheckDataMatch);
            testQc.Tolerance = 1;
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a number column that is not being checked for data match yet has a non-zero tolerance.");
            rrController.UseSampleData();
            testRecon = rrController.Recons.ReconReports.Find(recon => recon.Columns.Exists(col => col.Type == ColumnType.number && !col.CheckDataMatch));
            testQc = testRecon.Columns.Find(col => col.Type == ColumnType.number && col.CheckDataMatch && col.Tolerance > 0);
            testQc.Tolerance = -1 * testQc.Tolerance;
            Assert.IsFalse(rrController.ReadyToRun(), "Did not detect a negative tolerance.");
        }
        #endregion Test Validation Logic
    }
}
