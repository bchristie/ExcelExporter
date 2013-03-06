using System;
using System.Collections.Generic;

namespace ExcelExporter.Demo
{
    class ExampleObject
    {
        [ExcelColumn("Item #", ExcelColumnFormat.Text)]
        public String ItemNumber { get; set; }

        [ExcelColumn("Description", ExcelColumnFormat.Text)]
        public String Description { get; set; }

        [ExcelColumn("Qty", ExcelColumnFormat.Number)]
        public Int32 Quantity { get; set; }

        [ExcelColumn("Unit Price", ExcelColumnFormat.Currency)]
        public Double UnitPrice { get; set; }

        [ExcelColumn("Extended Price", ExcelColumnFormat.Currency)]
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
            List<ExampleObject> examples = new List<ExampleObject>();
            Random random = new Random(DateTime.Now.Millisecond);
            String[] itemNames = new[]{ "Apple", "Banana", "Cherry", "Grapefruit", "Kiwi", "Lime", "Mango", "Orange" };
            for (Int32 i = 0; i < itemNames.Length; i++)
            {
                examples.Add(new ExampleObject
                {
                    ItemNumber = (i + 1).ToString("D3"),
                    Description = itemNames[i],
                    Quantity = random.Next(1,10),
                    UnitPrice = random.NextDouble() * 10,
                    LastPurchased = DateTime.Now.AddDays(random.Next(-5,5))
                });
            }

            ExcelExporter exporter = new ExcelExporter();
            exporter.ExportToSheet<ExampleObject>(examples, "Examples");

            examples.Add(new ExampleObject { ItemNumber = "999", Description = "Last Object", Quantity = 99, UnitPrice = 999.99, LastPurchased = DateTime.Now.AddDays(99) });
            exporter.ExportToSheet<ExampleObject>(examples, "Examples (+1)");
        }
    }
}
