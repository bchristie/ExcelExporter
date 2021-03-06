ExcelExporter
=============

Simple Object to Excel exporter.

Synopsis
--------

I needed a way to take a single object and an Enumerable of objects and dumpt them in to excel. I also needed the
data to stay preserved in a way that Excel didn't ruin fields with its content detection. For example, under normal
circumstances exporting an object as a CSV that contains "00123" would be observed in Excel as simply the number
"123"; I just couldn't have that.

Solution
--------

To preserve information, this parser uses a coupling of reflection to know column titles (object properties become
titles) and an `ExcelColumnAttribute` decorator for further customization. e.g.

    class Invoice
    {
        [ExcelColumn("Invoice", ExcelColumnFormat.Text)]
        public String InvoiceId { get; set; }

        [ExcelColumn("Amount Due", ExcelColumnFormat.Currency)]
        public Double Amount { get; set; }
    }

Resulting in something similar to:

<kbd>Invoice</kbd><kbd>Amount Due</kbd>  
<kbd>00123&nbsp;</kbd><kbd>$1,024.99&nbsp;&emsp;</kbd>  
<kbd>00124&nbsp;</kbd><kbd>$2,048.99&nbsp;&emsp;</kbd>  
<kbd>00125&nbsp;</kbd><kbd>$4,096.99&nbsp;&emsp;</kbd>
