using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelExporter.Tests
{
    /// <summary>
    /// Class ExcelExporterTest
    /// </summary>
    [TestClass]
    public class BasicFunctionalityTest
    {
        private ExcelExporter exporter;

        [TestInitialize]
        public void Initialize()
        {
            this.exporter = new ExcelExporter();
        }

        [TestCleanup]
        public void Cleanup()
        {
            this.exporter.Dispose();
        }

        /// <summary>
        /// Cannot_export_null_objects
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Cannot_export_null_object()
        {
            this.exporter.ExportToSheet<Object>(item: null);
        }

        /// <summary>
        /// Cannot_export_null_collections
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Cannot_export_null_collection()
        {
            this.exporter.ExportToSheet<Object>(items: null);
        }

        /// <summary>
        /// Cannot_export_empty_collections
        /// </summary>
        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void Cannot_export_empty_collection()
        {
            this.exporter.ExportToSheet<Object>(items: new Object[0]);
        }
    }
}
