using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExporter.Xl
{
    /// <summary>
    /// Class Book
    /// </summary>
    internal class Book : IDisposable
    {
        /// <summary>
        /// The xl book
        /// </summary>
        private Excel.Workbook xlBook;

        #region Ctor

        /// <summary>
        /// Initializes a new instance of the <see cref="Book"/> class.
        /// </summary>
        /// <param name="xlApp">The xl app.</param>
        internal Book(Excel.Application xlApp)
        {
            if (xlApp == null)
            {
                throw new ArgumentNullException("xlApp");
            }

            this.xlBook = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            if (this.xlBook == null)
            {
                throw new ApplicationException("Unable to create workbook.");
            }
        }

        #endregion

        /// <summary>
        /// Adds the sheet.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>ExcelExport.Xl.Sheet.</returns>
        internal Sheet AddSheet(String sheetName = null)
        {
            return new Xl.Sheet(this.xlBook);
        }

        #region Dtor

        /// <summary>
        /// Finalizes an instance of the <see cref="Book"/> class.
        /// </summary>
        ~Book()
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

                }

                this.xlBook.Close();
                Marshal.FinalReleaseComObject(this.xlBook);
            }
        }
        #endregion
    }
}
