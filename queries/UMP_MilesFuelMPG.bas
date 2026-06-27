Operation =1
Option =0
Begin InputTables
    Name ="vw_UnitMileageProcessed"
End
Begin OutputColumns
    Alias ="UnitRef"
    Expression ="vw_UnitMileageProcessed.UnitId"
    Alias ="Mileage"
    Expression ="Max(vw_UnitMileageProcessed.MilesLTD)"
    Alias ="OverAllMPG"
    Expression ="IIf([Fuel]>0 And [Mileage]>0,Round([Mileage]/[Fuel],2),0)"
    Alias ="Fuel"
    Expression ="Max(vw_UnitMileageProcessed.FuelLTD)"
    Alias ="LastReport"
    Expression ="Max(vw_UnitMileageProcessed.LogActivityDate)"
End
Begin Groups
    Expression ="vw_UnitMileageProcessed.UnitId"
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
        dbText "Name" ="UnitRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OverAllMPG"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Mileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fuel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_UnitMileageProcessed.LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MinOfLogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastReport"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =275
    Top =201
    Right =1533
    Bottom =847
    Left =-1
    Top =-1
    Right =1226
    Bottom =334
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =420
        Top =0
        Right =736
        Bottom =359
        Top =0
        Name ="vw_UnitMileageProcessed"
        Name =""
    End
End
