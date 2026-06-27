Operation =1
Option =0
Where ="(((RemarketExpense.AssetId) In (SELECT [AssetId] FROM [RemarketExpense] As Tmp G"
    "ROUP BY [AssetId],[InvoiceNumber],[ExpenseTypeId] HAVING Count(*)>1  And [Invoic"
    "eNumber] = [RemarketExpense].[InvoiceNumber] And [ExpenseTypeId] = [RemarketExpe"
    "nse].[ExpenseTypeId])))"
Begin InputTables
    Name ="RemarketExpense"
    Name ="Vendors"
End
Begin OutputColumns
    Expression ="RemarketExpense.AssetId"
    Expression ="RemarketExpense.VendorId"
    Expression ="RemarketExpense.InvoiceNumber"
    Expression ="RemarketExpense.ExpenseTypeId"
    Expression ="RemarketExpense.RMKTExpenseId"
    Expression ="RemarketExpense.Expense"
    Expression ="RemarketExpense.ExpenseDate"
    Expression ="Vendors.VendorName"
End
Begin Joins
    LeftTable ="RemarketExpense"
    RightTable ="Vendors"
    Expression ="RemarketExpense.VendorId = Vendors.VendorID"
    Flag =1
End
Begin OrderBy
    Expression ="RemarketExpense.AssetId"
    Flag =0
    Expression ="RemarketExpense.VendorId"
    Flag =0
    Expression ="RemarketExpense.InvoiceNumber"
    Flag =0
    Expression ="RemarketExpense.ExpenseTypeId"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="([Find dup for RemarketExpense2].[InvoiceNumber]=\"82033\")"
dbMemo "OrderBy" ="[Find dup for RemarketExpense2].[InvoiceNumber], [Find dup for RemarketExpense2]"
    ".[AssetId]"
Begin
    Begin
        dbText "Name" ="RemarketExpense.AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketExpense.VendorId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketExpense.InvoiceNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketExpense.ExpenseTypeId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketExpense.RMKTExpenseId"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketExpense.Expense"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors.VendorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RemarketExpense.ExpenseDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =161
    Top =-14
    Right =1671
    Bottom =845
    Left =-1
    Top =-1
    Right =1484
    Bottom =273
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =55
        Top =13
        Right =220
        Bottom =178
        Top =0
        Name ="RemarketExpense"
        Name =""
    End
    Begin
        Left =414
        Top =77
        Right =579
        Bottom =242
        Top =0
        Name ="Vendors"
        Name =""
    End
End
