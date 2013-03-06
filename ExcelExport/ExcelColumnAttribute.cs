using System;

namespace ExcelExporter
{
    /// <summary>
    /// Enum ExcelColumnFormat
    /// </summary>
    public enum ExcelColumnFormat
    {
        /// <summary>
        /// General formatting (no special attention is paid to how it's displayed)
        /// </summary>
        General,

        /// <summary>
        /// Numeric formatting
        /// </summary>
        Number,

        /// <summary>
        /// currency formatting
        /// </summary>
        Currency,

        /// <summary>
        /// Accounting formatting
        /// </summary>
        Accounting,

        /// <summary>
        /// Short date formatting
        /// </summary>
        ShortDate,

        /// <summary>
        /// Long date formatting
        /// </summary>
        LongDate,

        /// <summary>
        /// Time formatting
        /// </summary>
        Time,

        /// <summary>
        /// Percentage formatting
        /// </summary>
        Percentage,

        /// <summary>
        /// Fraction formatting
        /// </summary>
        Fraction,

        /// <summary>
        /// Scientific formatting
        /// </summary>
        Scientific,

        /// <summary>
        /// Text formatting
        /// +
        /// </summary>
        Text
    }

    /// <summary>
    /// Enum ExcelColumnSummary
    /// </summary>
    public enum ExcelColumnSummary
    {
        /// <summary>
        /// Do not add a summary
        /// </summary>
        None,

        /// <summary>
        /// Display a total for the column
        /// </summary>
        Total,

        /// <summary>
        /// Display an average for the column
        /// </summary>
        Average
    }

    /// <summary>
    /// Property decorations providing more control over how it is exported to Excel.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// Gets the column heading.
        /// </summary>
        /// <value>The column heading.</value>
        public String Title { get; private set; }

        /// <summary>
        /// Gets the column format.
        /// </summary>
        /// <value>The column format.</value>
        public ExcelColumnFormat Format { get; private set; }

        /// <summary>
        /// Gets or sets the summary setting.
        /// </summary>
        /// <value>The summary.</value>
        public ExcelColumnSummary Summary { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelColumnAttribute"/> class.
        /// </summary>
        /// <param name="columnTitle">The column heading.</param>
        /// <param name="columnFormat">The column format.</param>
        public ExcelColumnAttribute(String columnTitle, ExcelColumnFormat columnFormat = ExcelColumnFormat.General)
        {
            this.Title = columnTitle ?? String.Empty;
            this.Format = columnFormat;
        }
    }
}
