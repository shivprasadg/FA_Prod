Operation =1
Option =0
Begin InputTables
    Name ="RemarketExpense"
End
Begin OutputColumns
    Expression ="RemarketExpense.AssetId"
    Alias ="Title"
    Expression ="Sum(CCur(IIf([ExpenseTypeId]=1,[Expense],0)))"
    Alias ="Transport"
    Expression ="Sum(CCur(IIf([ExpenseTypeId]=2,[Expense],0)))"
    Alias ="Advertising"
    Expression ="Sum(CCur(IIf([ExpenseTypeId]=4,[Expense],0)))"
    Alias ="Repairs"
    Expression ="Sum(CCur(IIf([ExpenseTypeId]=5,[Expense],0)))"
    Alias ="Paint"
    Expression ="Sum(CCur(IIf([ExpenseTypeId]=6,[Expense],0)))"
    Alias ="Bankshare"
    Expression ="Sum(CCur(IIf([ExpenseTypeId]=7,[Expense],0)))"
    Alias ="Interest"
    Expression ="Sum(CCur(IIf([ExpenseTypeId]=8,[Expense],0)))"
    Alias ="Commission"
    Expression ="Sum(CCur(IIf([ExpenseTypeId]=9,[Expense],0)))"
    Alias ="Other"
    Expression ="Sum(CCur(IIf([ExpenseTypeId]=10,[Expense],0)))"
    Alias ="SubTotalExpense"
    Expression ="[Title]+[Transport]+[Advertising]+[Repairs]+[Paint]+[Bankshare]+[Interest]+[Comm"
        "ission]+[Other]"
End
Begin Groups
    Expression ="RemarketExpense.AssetId"
    GroupLevel =0
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
        dbText "Name" ="RemarketExpense.AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Repairs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transport"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Advertising"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Paint"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bankshare"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Commission"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Other"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Interest"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubTotalExpense"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-2
    Top =19
    Right =1464
    Bottom =696
    Left =-1
    Top =-1
    Right =1434
    Bottom =343
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =53
        Top =45
        Right =416
        Bottom =319
        Top =0
        Name ="RemarketExpense"
        Name =""
    End
End
