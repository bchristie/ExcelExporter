using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExporter.Xl
{
    /// <summary>
    /// Class Sheet
    /// </summary>
    internal class Sheet : IDisposable
    {
        /// <summary>
        /// The xl sheet
        /// </summary>
        private Excel.Worksheet xlSheet;

        #region Ctor

        /// <summary>
        /// Initializes a new instance of the <see cref="Sheet"/> class.
        /// </summary>
        /// <param name="xlBook">The xl book.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        internal Sheet(Excel.Workbook xlBook, String sheetName = null)
        {
            this.xlSheet = xlBook.Worksheets.Add();
            if (!String.IsNullOrEmpty(sheetName))
            {
                this.xlSheet.Name = sheetName;
            }
        }

        #endregion

        /// <summary>
        /// Loads the headers.
        /// </summary>
        /// <typeparam name="T">The type of the object.</typeparam>
        /// <exception cref="System.NotImplementedException"></exception>
        internal void LoadHeaders<T>() where T : class
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));

        }

        #region Dtor

        /// <summary>
        /// Finalizes an instance of the <see cref="Sheet"/> class.
        /// </summary>
        ~Sheet()
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

                Marshal.ReleaseComObject(this.xlSheet);
            }
        }
        #endregion
    }
}
