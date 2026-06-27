Operation =1
Option =0
Begin InputTables
    Name ="UnitsByYear"
End
Begin OutputColumns
    Expression ="UnitsByYear.DYear"
    Expression ="UnitsByYear.UnitType"
    Alias ="Active"
    Expression ="Sum(([A2014]+[A2015]+[A2016]+[A2017]+[A2018]))"
    Alias ="New"
    Expression ="Sum(([New2014]+[New2015]+[New2016]+[New2017]+[New2018]))"
    Alias ="Added"
    Expression ="Sum(([added2014]+[added2015]+[added2016]+[added2017]+[added2018]))"
End
Begin OrderBy
    Expression ="UnitsByYear.DYear"
    Flag =0
    Expression ="UnitsByYear.UnitType"
    Flag =0
End
Begin Groups
    Expression ="UnitsByYear.DYear"
    GroupLevel =0
    Expression ="UnitsByYear.UnitType"
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
        dbText "Name" ="UnitsByYear.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Active"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="New"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsByYear.DYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Added"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =187
    Top =43
    Right =1511
    Bottom =972
    Left =-1
    Top =-1
    Right =1306
    Bottom =652
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =131
        Top =45
        Right =509
        Bottom =546
        Top =0
        Name ="UnitsByYear"
        Name =""
    End
End
