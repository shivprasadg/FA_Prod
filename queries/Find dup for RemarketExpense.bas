Operation =1
Option =0
Where ="(((RemarketExpense.[AssetId]) In (SELECT [AssetId] FROM [RemarketExpense] As Tmp"
    " GROUP BY [AssetId],[VendorId],[InvoiceNumber],[ExpenseTypeId] HAVING Count(*)>1"
    "  And [VendorId] = [RemarketExpense].[VendorId] And [InvoiceNumber] = [RemarketE"
    "xpense].[InvoiceNumber] And [ExpenseTypeId] = [RemarketExpense].[ExpenseTypeId])"
    "))"
Begin InputTables
    Name ="RemarketExpense"
End
Begin OutputColumns
    Expression ="RemarketExpense.[AssetId]"
    Expression ="RemarketExpense.[VendorId]"
    Expression ="RemarketExpense.[InvoiceNumber]"
    Expression ="RemarketExpense.[ExpenseTypeId]"
    Expression ="RemarketExpense.[RMKTExpenseId]"
    Expression ="RemarketExpense.[Expense]"
    Expression ="RemarketExpense.[ExpenseDate]"
End
Begin OrderBy
    Expression ="RemarketExpense.[AssetId]"
    Flag =0
    Expression ="RemarketExpense.[VendorId]"
    Flag =0
    Expression ="RemarketExpense.[InvoiceNumber]"
    Flag =0
    Expression ="RemarketExpense.[ExpenseTypeId]"
    Flag =0
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
        dbText "Name" ="[Expense]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[AssetId]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[ExpenseDate]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[VendorId]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[InvoiceNumber]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[ExpenseTypeId]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[RMKTExpenseId]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1833"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =60
    Top =1
    Right =1570
    Bottom =1089
    Left =-1
    Top =-1
    Right =1479
    Bottom =311
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
End
