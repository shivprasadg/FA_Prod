Operation =1
Option =2
Begin InputTables
    Name ="Unit_Vin_Remarket_Temp"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Expression ="Unit_Vin_Remarket_Temp.Asset"
    Expression ="Units.UnitUnitNum"
    Expression ="Unit_Vin_Remarket_Temp.Unit"
    Expression ="Unit_Vin_Remarket_Temp.VIN"
    Expression ="Units.UnitVIN"
    Expression ="Unit_Vin_Remarket_Temp.Make"
    Expression ="Unit_Vin_Remarket_Temp.State"
    Alias ="Match"
    Expression ="IIf(Trim([VIN])=Trim([UnitVIN]),\"Yes\",\"No\")"
End
Begin Joins
    LeftTable ="Units"
    RightTable ="Unit_Vin_Remarket_Temp"
    Expression ="Units.UnitUnitNum = Unit_Vin_Remarket_Temp.Unit"
    Flag =1
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
dbMemo "Filter" ="([VINImport_Remarketing].[Match]<>\"Yes\" Or [VINImport_Remarketing].[Match] IS "
    "Null)"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitVIN"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit_Vin_Remarket_Temp.VIN"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit_Vin_Remarket_Temp.Make"
        dbInteger "ColumnWidth" ="2880"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit_Vin_Remarket_Temp.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit_Vin_Remarket_Temp.Unit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Match"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit_Vin_Remarket_Temp.Asset"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =254
    Top =10
    Right =1866
    Bottom =842
    Left =-1
    Top =-1
    Right =1580
    Bottom =481
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =304
        Top =66
        Right =470
        Bottom =429
        Top =0
        Name ="Unit_Vin_Remarket_Temp"
        Name =""
    End
    Begin
        Left =59
        Top =26
        Right =203
        Bottom =550
        Top =0
        Name ="Units"
        Name =""
    End
End
