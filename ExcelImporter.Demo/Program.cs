using System;
using System.Collections.Generic;

namespace ExcelExporter.Demo
{
    class SimpleExample
    {
        [ExcelColumn("Item #", ExcelColumnFormat.Text)]
        public String ItemNumber { get; set; }

        [ExcelColumn("Description", ExcelColumnFormat.Text)]
        public String Description { get; set; }

        [ExcelColumn("Qty", ExcelColumnFormat.Number)]
        public Int32 Quantity { get; set; }

        [ExcelColumn("Unit Price", ExcelColumnFormat.Currency, Summary = ExcelColumnSummary.Average)]
        public Double UnitPrice { get; set; }

        [ExcelColumn("Extended Price", ExcelColumnFormat.Currency, Summary = ExcelColumnSummary.Total)]
        public Double ExtendedPrice
        {
            get { return this.Quantity * this.UnitPrice; }
        }

        [ExcelColumn("Last Purchase Date", ExcelColumnFormat.ShortDate)]
        public DateTime LastPurchased { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            List<SimpleExample> simpleExamples = new List<SimpleExample>();
            Random random = new Random(DateTime.Now.Millisecond);
            String[] itemNames = new[]{ "Apple", "Banana", "Cherry", "Grapefruit", "Kiwi", "Lime", "Mango", "Orange" };
            for (Int32 i = 0; i < itemNames.Length; i++)
            {
                simpleExamples.Add(new SimpleExample
                {
                    ItemNumber = (i + 1).ToString("D3"),
                    Description = itemNames[i],
                    Quantity = random.Next(1,10),
                    UnitPrice = random.NextDouble() * 10,
                    LastPurchased = DateTime.Now.AddDays(random.Next(-5,5))
                });
            }

            using (ExcelExporter exporter = new ExcelExporter())
            {
                exporter.ExportToSheet<SimpleExample>(simpleExamples, "Examples");
            }
        }
    }
}
