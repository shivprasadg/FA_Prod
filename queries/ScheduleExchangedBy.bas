Operation =1
Option =0
Begin InputTables
    Name ="ExitOptionsMatchedExchangeIT"
End
Begin OutputColumns
    Expression ="ExitOptionsMatchedExchangeIT.SchID"
    Alias ="NewSchGrp"
    Expression ="Max([NewSchedule] & \" | \" & [NewGroup])"
    Alias ="ReplacedByScheduleID"
    Expression ="First(ExitOptionsMatchedExchangeIT.ReplacedByScheduleID)"
    Alias ="ReplacedByGroupID"
    Expression ="First(ExitOptionsMatchedExchangeIT.ReplacedByGroupID)"
End
Begin Groups
    Expression ="ExitOptionsMatchedExchangeIT.SchID"
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
        dbText "Name" ="ExitOptionsMatchedExchangeIT.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReplacedByScheduleID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReplacedByGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewSchGrp"
        dbInteger "ColumnWidth" ="4815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =237
    Top =19
    Right =1830
    Bottom =900
    Left =-1
    Top =-1
    Right =1561
    Bottom =546
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =122
        Top =79
        Right =581
        Bottom =472
        Top =0
        Name ="ExitOptionsMatchedExchangeIT"
        Name =""
    End
End
