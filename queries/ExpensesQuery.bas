Operation =1
Option =0
Begin InputTables
    Name ="Expenses"
End
Begin OutputColumns
    Expression ="Expenses.ExpenseDate"
    Expression ="Expenses.ExpenseTypeRef"
    Expression ="Expenses.ExpenseID"
    Expression ="Expenses.DealsRef"
    Expression ="Expenses.GroupsRef"
    Expression ="Expenses.ExpenseRep"
    Expression ="Expenses.ExpenseDesc"
    Expression ="Expenses.ExpenseMemo"
    Expression ="Expenses.ExpenseAmt"
End
Begin OrderBy
    Expression ="Expenses.ExpenseDate"
    Flag =1
    Expression ="Expenses.ExpenseTypeRef"
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
        dbText "Name" ="Expenses.ExpenseDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expenses.ExpenseTypeRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expenses.ExpenseID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =412
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =221
        Top =0
        Name ="Expenses"
        Name =""
    End
End
