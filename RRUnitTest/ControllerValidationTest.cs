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
            Assert.IsTrue(rrController.ReadyToRun(), string.Join(",", rrController.GetValidationErrors()));           
        }

        [TestMethod]
        public void DetectDupTemplateNames()
        {
            rrController.UseSampleData();
            rrController.Sources.ConnStringTemplates[0].Name = rrController.Sources.ConnStringTemplates[1].Name;
            Assert.IsFalse(rrController.ReadyToRun());
        }

        [TestMethod]
        public void DetectUnknownTemplate()
        {
            rrController.UseSampleData();
            rrController.Sources.ConnectionStrings[0].TemplateName += "XYZ";
            Assert.IsFalse(rrController.ReadyToRun());
        }

        [TestMethod]
        public void DetectConnStringMissingPlaceholder()
        {

        }
        #endregion Test Validation Logic
    }
}
