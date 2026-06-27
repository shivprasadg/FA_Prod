Operation =1
Option =0
Where ="(((RmktInvoiceRequestTbl.InvoiceRequestID) Is Not Null))"
Begin InputTables
    Name ="RmktInvoiceRequestTbl"
End
Begin OutputColumns
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestID"
    Expression ="RmktInvoiceRequestTbl.InvoiceRequestDate"
    Expression ="RmktInvoiceRequestTbl.InvoicePaidDate"
    Expression ="RmktInvoiceRequestTbl.BuyerID"
    Alias ="Status"
    Expression ="Switch(Not IsDate([RmktInvoiceRequestTbl].[InvoiceRequestDate]),\"Open\",IsDate("
        "[RmktInvoiceRequestTbl].[InvoiceRequestDate]) And Not IsDate([RmktInvoiceRequest"
        "Tbl].[InvoicePaidDate]) And Not IsDate([RmktInvoiceRequestTbl].[InvoiceSentDate]"
        "),\"Request\",Not IsDate([RmktInvoiceRequestTbl].[InvoicePaidDate]) And IsDate(["
        "RmktInvoiceRequestTbl].[InvoiceSentDate]),\"Sent\",IsDate([RmktInvoiceRequestTbl"
        "].[InvoicePaidDate]),\"Paid\")"
    Expression ="RmktInvoiceRequestTbl.InvoiceNumber"
    Expression ="RmktInvoiceRequestTbl.InvoiceDate"
    Expression ="RmktInvoiceRequestTbl.InvoiceDueDate"
    Expression ="RmktInvoiceRequestTbl.InvoiceSentDate"
    Expression ="RmktInvoiceRequestTbl.InvoiceLink"
    Expression ="RmktInvoiceRequestTbl.TaxRateToUse"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceRequestID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.BuyerID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.TaxRateToUse"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3465"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceDueDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceLink"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceSentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmktInvoiceRequestTbl.InvoiceRequestDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-2
    Top =77
    Right =1436
    Bottom =1005
    Left =-1
    Top =-1
    Right =1414
    Bottom =440
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =85
        Top =40
        Right =470
        Bottom =387
        Top =0
        Name ="RmktInvoiceRequestTbl"
        Name =""
    End
End
