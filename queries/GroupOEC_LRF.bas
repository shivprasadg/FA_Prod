Operation =1
Option =0
Begin InputTables
    Name ="UnitOEC2"
End
Begin OutputColumns
    Alias ="GroupId"
    Expression ="UnitOEC2.GroupREF"
    Alias ="OEC"
    Expression ="Sum(UnitOEC2.TotalCost)"
    Alias ="GroupLoRent"
    Expression ="Sum(UnitOEC2.LORent)"
    Alias ="GroupSyndRent"
    Expression ="Sum(UnitOEC2.SYNDRent)"
    Alias ="GroupLOLRF"
    Expression ="GetLFR([GroupLORent],[OEC])"
    Alias ="GroupSyndLRF"
    Expression ="GetLFR([GroupSyndRent],[OEC])"
End
Begin Groups
    Expression ="UnitOEC2.GroupREF"
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
        dbText "Name" ="GroupSyndRent"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupLoRent"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupLOLRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupSyndLRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1377
    Bottom =974
    Left =-1
    Top =-1
    Right =1345
    Bottom =617
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =82
        Top =77
        Right =682
        Bottom =627
        Top =0
        Name ="UnitOEC2"
        Name =""
    End
End
