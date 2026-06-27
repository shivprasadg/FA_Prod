dbMemo "SQL" ="SELECT Units.UnitID, Vendors.VendorName, RemarketExpenseType.RemarketExpense, Rm"
    "kExpense.InvAmt, RmkExpense.InvoiceNumber, RmkExpense.InvoiceDate\015\012FROM (("
    "RmkExpense RIGHT JOIN (Units INNER JOIN Rmkt ON Units.UnitID=Rmkt.UnitRef) ON Rm"
    "kExpense.RmktRef=Rmkt.RmktId) LEFT JOIN Vendors ON RmkExpense.VendorRef=Vendors."
    "VendorID) LEFT JOIN RemarketExpenseType ON RmkExpense.ExpenseTypeId=RemarketExpe"
    "nseType.RemarketExpenseTypeId;\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketExpenseType.RemarketExpense"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmkExpense.InvAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmkExpense.InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RmkExpense.InvoiceDt"
        dbLong "AggregateType" ="-1"
    End
End
