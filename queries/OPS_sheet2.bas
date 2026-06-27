Operation =1
Option =0
Begin InputTables
    Name ="ScheduleUnitDocs4"
    Name ="Cost_Refrig"
    Name ="Units"
    Name ="Last Title Entry"
End
Begin OutputColumns
    Expression ="ScheduleUnitDocs4.Lessee"
    Expression ="ScheduleUnitDocs4.SchNo"
    Expression ="ScheduleUnitDocs4.[Lease No]"
    Expression ="ScheduleUnitDocs4.Group"
    Expression ="ScheduleUnitDocs4.[Unit No]"
    Expression ="ScheduleUnitDocs4.AssetId"
    Expression ="ScheduleUnitDocs4.Status"
    Expression ="ScheduleUnitDocs4.Unitvin"
    Expression ="ScheduleUnitDocs4.Desc"
    Expression ="ScheduleUnitDocs4.Location"
    Expression ="ScheduleUnitDocs4.StreetAddress"
    Expression ="ScheduleUnitDocs4.City"
    Expression ="ScheduleUnitDocs4.State"
    Expression ="ScheduleUnitDocs4.Zip"
    Expression ="Cost_Refrig.RefrigMake"
    Alias ="RefrigCost"
    Expression ="CCur(Nz([Cost_Refrig].[Refrig],0))"
    Expression ="Cost_Refrig.RefrigModel"
    Expression ="ScheduleUnitDocs4.RefrigSerial"
    Alias ="PlateNum"
    Expression ="Units.UnitPlateNum"
    Alias ="InServiceInd"
    Expression ="Units.UnitInServInd"
    Alias ="InServiceDate"
    Expression ="ScheduleUnitDocs4.InService"
    Alias ="MSORecv"
    Expression ="Units.UnitMSORecd"
    Alias ="MSOCopy"
    Expression ="Units.UnitMSOCopy"
    Alias ="TitleAppSent"
    Expression ="Units.UnitTASent"
    Alias ="LastTitleLog"
    Expression ="[Last Title Entry].LogDate"
    Alias ="TitleAction"
    Expression ="[Last Title Entry].Action"
    Expression ="ScheduleUnitDocs4.UnitCost"
    Alias ="ChildPartCost"
    Expression ="[TotalCost]-[UnitCost]"
    Alias ="Expr1"
    Expression ="ScheduleUnitDocs4.TotalCost"
End
Begin Joins
    LeftTable ="ScheduleUnitDocs4"
    RightTable ="Cost_Refrig"
    Expression ="ScheduleUnitDocs4.AssetId = Cost_Refrig.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs4"
    RightTable ="Units"
    Expression ="ScheduleUnitDocs4.AssetId = Units.UnitID"
    Flag =2
    LeftTable ="ScheduleUnitDocs4"
    RightTable ="Last Title Entry"
    Expression ="ScheduleUnitDocs4.AssetId = [Last Title Entry].AssetID"
    Flag =2
End
Begin OrderBy
    Expression ="ScheduleUnitDocs4.Group"
    Flag =0
    Expression ="ScheduleUnitDocs4.[Unit No]"
    Flag =0
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
        dbText "Name" ="ScheduleUnitDocs4.[Lease No]"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.Lessee"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.Group"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.[Unit No]"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.AssetId"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.Unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.SchNo"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.RefrigSerial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.Status"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.StreetAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InServiceInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MSORecv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="titleAppSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleAction"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs4.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleLog"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MSOCopy"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =90
    Top =71
    Right =1637
    Bottom =937
    Left =-1
    Top =-1
    Right =1515
    Bottom =458
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =92
        Top =85
        Right =280
        Bottom =561
        Top =0
        Name ="ScheduleUnitDocs4"
        Name =""
    End
    Begin
        Left =314
        Top =93
        Right =486
        Bottom =304
        Top =0
        Name ="Cost_Refrig"
        Name =""
    End
    Begin
        Left =530
        Top =96
        Right =739
        Bottom =578
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =791
        Top =136
        Right =955
        Bottom =311
        Top =0
        Name ="Last Title Entry"
        Name =""
    End
End
