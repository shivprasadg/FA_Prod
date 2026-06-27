Operation =1
Option =0
Where ="(((BillingRules.RuleType)=\"FeeLabel\"))"
Begin InputTables
    Name ="BillingRules"
End
Begin OutputColumns
    Expression ="BillingRules.RuleType"
    Expression ="BillingRules.RuleName"
    Expression ="BillingRules.RuleString"
    Expression ="BillingRules.RuleValue"
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
        dbText "Name" ="BillingRules.RuleType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillingRules.RuleName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillingRules.RuleString"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillingRules.RuleValue"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =236
    Top =123
    Right =1699
    Bottom =951
    Left =-1
    Top =-1
    Right =1431
    Bottom =528
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =135
        Top =36
        Right =279
        Bottom =180
        Top =0
        Name ="BillingRules"
        Name =""
    End
End
