Operation =4
Option =0
Where ="(((Units.UnitStatus)=\"T\") AND ((Units.unitoffleasedt) Is Not Null And (Units.u"
    "nitoffleasedt)<Date()-30))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Name ="Units.UnitStatus"
    Expression ="\"O\""
    Name ="Units.UnitOLInd"
    Expression ="\"A\""
    Name ="Units.ReportEndDate"
    Expression ="[unitoffleasedt]"
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ReportStartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.ReportEndDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitoffleasedt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitOLInd"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =57
    Top =2
    Right =1484
    Bottom =1000
    Left =-1
    Top =-1
    Right =1395
    Bottom =697
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =447
        Bottom =651
        Top =0
        Name ="Units"
        Name =""
    End
End
