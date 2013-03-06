using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExporter
{
    /// <summary>
    /// Class to export various objects and collections of objects to an excel file.
    /// </summary>
    public class ExcelExporter : IDisposable
    {
        /// <summary>
        /// The xl app
        /// </summary>
        private Excel.Application xlApp;

        /// <summary>
        /// The xl book
        /// </summary>
        private Excel.Workbook xlBook;

        /// <summary>
        /// The title bar height
        /// </summary>
        private Int32 titleHeight = 1;

        #region Ctor

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelExporter"/> class.
        /// </summary>
        /// <exception cref="System.ApplicationException">
        /// Unable to connect to Excel.
        /// or
        /// Unable to create a new workbook.
        /// </exception>
        public ExcelExporter()
        {
            this.xlApp = new Excel.Application() as Excel.Application;
            if (this.xlApp == null)
            {
                throw new ApplicationException("Unable to connect to Excel.");
            }

            this.xlBook = this.xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet) as Excel.Workbook;
            if (this.xlBook == null)
            {
                throw new ApplicationException("Unable to create a new workbook.");
            }

            //this.xlBook = new Xl.Book(this.xlApp);
            //if (this.xlBook == null)
            //{
            //    throw new ApplicationException("Unable to create a new workbook.");
            //}
        }

        #endregion

        /// <summary>
        /// Exports the specified item.
        /// </summary>
        /// <typeparam name="T">Type of the item to export.</typeparam>
        /// <param name="item">The item.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <exception cref="System.ArgumentNullException">item;item cannot be null.</exception>
        public void ExportToSheet<T>(T item, String sheetName = null) where T : class
        {
            if (item == null)
            {
                throw new ArgumentNullException("item", "item cannot be null.");
            }
            this.ExportToSheet<T>(new[] { item }, sheetName);
        }

        /// <summary>
        /// Exports the specified items.
        /// </summary>
        /// <typeparam name="T">Type of the item to export.</typeparam>
        /// <param name="items">The items.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <exception cref="System.ArgumentNullException">items;items cannot be null or empty.</exception>
        /// <exception cref="System.ApplicationException">Unable to add new worksheet.</exception>
        public void ExportToSheet<T>(IEnumerable<T> items, String sheetName = null) where T : class
        {
            if (items == null || items.Count() == 0)
            {
                throw new ArgumentNullException("items", "items cannot be null or empty.");
            }

            // Excel always provides one empty sheet with a new workbook. Attempt to use this
            // sheet first before going to create a new one.
            Excel.Worksheet xlSheet = this.xlBook.ActiveSheet as Excel.Worksheet;
            if (xlSheet == null)
            {
                throw new ApplicationException("Unable to retrieve current worksheet.");
            }
            Excel.Range usedRange = xlSheet.UsedRange as Excel.Range;
            if (usedRange.Count > 1)
            {
                Excel.Sheets xlSheets = this.xlBook.Sheets as Excel.Sheets;
                if (xlSheets == null)
                {
                    throw new ApplicationException("Unable to fetch sheets to add new worksheet.");
                }
                xlSheet = xlSheets.Add() as Excel.Worksheet;
                Marshal.FinalReleaseComObject(xlSheets);
            }
            if (xlSheet == null)
            {
                throw new ApplicationException("Unable to add new worksheet.");
            }
            if (!String.IsNullOrEmpty(sheetName))
            {
                xlSheet.Name = sheetName;
            }

            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));

            this.AddTitleToSheet(xlSheet, properties);
            this.AddItemsToSheet<T>(xlSheet, properties, items);
            this.AddFinalTouches(xlSheet);

            Marshal.FinalReleaseComObject(usedRange);
            Marshal.FinalReleaseComObject(xlSheet);
        }

        /// <summary>
        /// Adds the items.
        /// </summary>
        /// <typeparam name="T">Type of the item to add.</typeparam>
        /// <param name="xlSheet">The xl sheet.</param>
        /// <param name="properties">The properties.</param>
        /// <param name="items">The items.</param>
        private void AddItemsToSheet<T>(Excel.Worksheet xlSheet, PropertyDescriptorCollection properties, IEnumerable<T> items)
        {
            T[] itemsArray = items.ToArray<T>();

            Int32 rows = itemsArray.Length;
            Int32 columns = properties.Count;

            Object[,] dataset = new Object[rows, columns];
            for (Int32 c = 0; c < columns; c++)
            {
                PropertyDescriptor property = properties[c];
                for (Int32 r = 0; r < rows; r++)
                {
                    dataset[r, c] = property.GetValue(itemsArray[r]);
                }
            }

            Excel.Range topLeft = xlSheet.Cells[1 + this.titleHeight, 1] as Excel.Range;
            Excel.Range bottomRight = xlSheet.Cells[this.titleHeight + rows, columns] as Excel.Range;
            Excel.Range inputRange = xlSheet.Range[topLeft, bottomRight] as Excel.Range;
            inputRange.Value2 = dataset;

            Marshal.FinalReleaseComObject(topLeft);
            Marshal.FinalReleaseComObject(bottomRight);
            Marshal.FinalReleaseComObject(inputRange);
        }

        private void AddFinalTouches(Excel.Worksheet xlSheet)
        {
            Excel.Range firstCell = xlSheet.Cells[1, 1] as Excel.Range;
            {
                Excel.Range titleRow = firstCell.EntireRow as Excel.Range;
                titleRow.RowHeight *= 2;
                titleRow.AutoFilter(Field: 1, Operator: Excel.XlAutoFilterOperator.xlAnd);
                titleRow.AutoFit();
                Marshal.ReleaseComObject(titleRow);

                Excel.Window xlWindow = this.xlApp.ActiveWindow;
                xlWindow.SplitRow = 1;
                xlWindow.FreezePanes = true;
                Marshal.FinalReleaseComObject(xlWindow);
            }
            Marshal.FinalReleaseComObject(firstCell);
        }

        /// <summary>
        /// Adds the title to the sheet.
        /// </summary>
        /// <param name="xlSheet">The xl sheet.</param>
        /// <param name="properties">The properties.</param>
        private void AddTitleToSheet(Excel.Worksheet xlSheet, PropertyDescriptorCollection properties)
        {
            Excel.Range titleCell = null;
            for (Int32 p = 0; p < properties.Count; p++)
            {
                PropertyDescriptor property = properties[p];

                String titleText = property.Name;
                ExcelColumnFormat? titleFormat = null;
                ExcelColumnSummary? titleSummary = null;

                ExcelColumnAttribute xlColAttr = property.Attributes.OfType<ExcelColumnAttribute>().FirstOrDefault();
                if (xlColAttr != null)
                {
                    titleText = xlColAttr.Title;
                    titleFormat = xlColAttr.Format;
                    titleSummary = xlColAttr.Summary;
                }

                titleCell = xlSheet.Cells[1, p + 1] as Excel.Range;
                titleCell.set_Value(Excel.XlRangeValueDataType.xlRangeValueDefault, titleText);
                titleCell.Style = "Heading 3";

                if (titleFormat.HasValue)
                {
                    Excel.Range titleColumn = titleCell.EntireColumn;
                    ApplyFormatToRange(titleColumn, titleFormat.Value);
                    Marshal.FinalReleaseComObject(titleColumn);
                }
                titleCell.NumberFormat = "@";

                if (titleSummary.HasValue)
                {
                    Boolean newRowSuccess = (Boolean)xlSheet.Rows.Insert(2, 1);

                    if (newRowSuccess)
                    {
                        this.titleHeight += 2;
                    }

                    //Marshal.FinalReleaseComObject(titleRow);
                }
            }
            Marshal.FinalReleaseComObject(titleCell);
        }

        /// <summary>
        /// Applies the format to range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="format">The format.</param>
        private static void ApplyFormatToRange(Excel.Range range, ExcelColumnFormat format)
        {
            switch (format)
            {
                case ExcelColumnFormat.Accounting:
                    range.NumberFormat = @"_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)";
                    break;
                case ExcelColumnFormat.Currency:
                    range.NumberFormat = "$#,##0.00";
                    break;
                case ExcelColumnFormat.Fraction:
                    range.NumberFormat = "# ?/?";
                    break;
                case ExcelColumnFormat.LongDate:
                    range.NumberFormat = "[$-F800]dddd, mmmm dd, yyyy";
                    break;
                case ExcelColumnFormat.Number:
                    range.NumberFormat = "0.00";
                    break;
                case ExcelColumnFormat.Percentage:
                    range.NumberFormat = "0.00%";
                    break;
                case ExcelColumnFormat.Scientific:
                    range.NumberFormat = "0.00E+00";
                    break;
                case ExcelColumnFormat.ShortDate:
                    range.NumberFormat = "m/d/yyyy";
                    break;
                case ExcelColumnFormat.Text:
                    range.NumberFormat = "@";
                    break;
                case ExcelColumnFormat.Time:
                    range.NumberFormat = "[$-F400]h:mm:ss AM/PM";
                    break;
                case ExcelColumnFormat.General:
                default:
                    range.NumberFormat = "General";
                    break;
            }
        }

        #region Dtor

        /// <summary>
        /// Finalizes an instance of the <see cref="ExcelExporter"/> class.
        /// </summary>
        ~ExcelExporter()
        {
            this.Dispose(false);
        }

        #endregion

        #region IDisposable

        /// <summary>
        /// Flag if class has been disposed
        /// </summary>
        private Boolean disposed;

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            this.xlApp.Visible = true;

            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        /// <param name="disposing">Flag if the method was called as IDisposable</param>
        private void Dispose(Boolean disposing)
        {
            if (!this.disposed)
            {
                this.disposed = true;

                if (disposing)
                {
                    //this.xlBook.Close();
                    //this.xlApp.Quit();
                }

                Marshal.FinalReleaseComObject(this.xlBook);
                Marshal.FinalReleaseComObject(this.xlApp);
            }
        }

        #endregion
    }
}
