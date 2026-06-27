Operation =1
Option =0
Begin InputTables
    Name ="UnfI_Non_Reporting"
    Name ="FleetStudyImportBase"
End
Begin OutputColumns
    Expression ="UnfI_Non_Reporting.CompanyName"
    Alias ="AssetID"
    Expression ="UnfI_Non_Reporting.UnitID"
    Alias ="UNFI_UnitNum"
    Expression ="FleetStudyImportBase.UnitNumber"
    Expression ="FleetStudyImportBase.VIN"
    Expression ="FleetStudyImportBase.UnitYear"
    Expression ="FleetStudyImportBase.UnitMake"
    Expression ="FleetStudyImportBase.UnitModel"
    Expression ="FleetStudyImportBase.UnitType"
    Alias ="FA_VIN"
    Expression ="UnfI_Non_Reporting.UnitVIN"
    Alias ="FA_UnitNum"
    Expression ="UnfI_Non_Reporting.OBCUnitNumber"
End
Begin Joins
    LeftTable ="UnfI_Non_Reporting"
    RightTable ="FleetStudyImportBase"
    Expression ="UnfI_Non_Reporting.UnitX = FleetStudyImportBase.UnitNumber"
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
Begin
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="FA_VIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.VIN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="FA_UnitNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="UnfI_Non_Reporting.CompanyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="FleetStudyImportBase.UnitModel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="UNFI_UnitNum"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnfI_Non_Reporting.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =201
    Top =55
    Right =1750
    Bottom =989
    Left =-1
    Top =-1
    Right =1517
    Bottom =616
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =211
        Top =107
        Right =355
        Bottom =435
        Top =0
        Name ="UnfI_Non_Reporting"
        Name =""
    End
    Begin
        Left =440
        Top =177
        Right =584
        Bottom =494
        Top =0
        Name ="FleetStudyImportBase"
        Name =""
    End
End
