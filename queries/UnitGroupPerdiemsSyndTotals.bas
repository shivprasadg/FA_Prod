Operation =1
Option =0
Begin InputTables
    Name ="UnitGroupPerdiemsSyndReport"
End
Begin OutputColumns
    Expression ="UnitGroupPerdiemsSyndReport.GroupID"
    Expression ="UnitGroupPerdiemsSyndReport.PDType"
    Alias ="PerdiemRate"
    Expression ="Avg(UnitGroupPerdiemsSyndReport.PerdiemRate)"
    Alias ="BilledPDDays"
    Expression ="Sum(IIf([PerdiemDays]<=0,0,[PerdiemDays]))"
    Alias ="PerdiemTotal"
    Expression ="Sum(UnitGroupPerdiemsSyndReport.PerdiemTotal)"
End
Begin Groups
    Expression ="UnitGroupPerdiemsSyndReport.GroupID"
    GroupLevel =0
    Expression ="UnitGroupPerdiemsSyndReport.PDType"
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
        dbText "Name" ="UnitGroupPerdiemsSyndReport.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BilledPDDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitGroupPerdiemsSyndReport.PDType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerdiemTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerdiemRate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1373
    Bottom =974
    Left =-1
    Top =-1
    Right =1341
    Bottom =617
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =211
        Top =57
        Right =547
        Bottom =286
        Top =0
        Name ="UnitGroupPerdiemsSyndReport"
        Name =""
    End
End
