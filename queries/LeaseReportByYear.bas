Operation =1
Option =0
Begin InputTables
    Name ="LeaseReport2"
End
Begin OutputColumns
    Expression ="LeaseReport2.LeaseYear"
    Expression ="LeaseReport2.Lessee"
    Expression ="LeaseReport2.[Lease No]"
    Expression ="LeaseReport2.Make"
    Expression ="LeaseReport2.UnitType"
    Alias ="UnitCount"
    Expression ="Count(LeaseReport2.AssetId)"
    Alias ="UnitCost"
    Expression ="Sum(LeaseReport2.UnitCost)"
    Alias ="Refrig"
    Expression ="Sum(LeaseReport2.Refrig)"
    Alias ="Decals"
    Expression ="Sum([MiscCost]+[CoNameCost]+[RegDecalCost])"
    Alias ="OBC"
    Expression ="Sum(LeaseReport2.OBC)"
    Alias ="OBCInstall"
    Expression ="Sum([LeaseReport2].[OBCInstall]+[OBCAdapt]+[OBCDome]+[OBCMulti])"
    Alias ="FeesTax"
    Expression ="Sum(LeaseReport2.FeesTax)"
    Alias ="Freight"
    Expression ="Sum(LeaseReport2.Freight)"
    Alias ="Warranty"
    Expression ="Sum(LeaseReport2.Warranty)"
    Alias ="Inspection"
    Expression ="Sum(LeaseReport2.Inspection)"
    Alias ="OilRig"
    Expression ="Sum(LeaseReport2.OilRig)"
    Alias ="TankBodies"
    Expression ="Sum(([TankBody]+[TankOther]))"
    Alias ="PTO"
    Expression ="Sum(LeaseReport2.PTO)"
    Alias ="OtherCost"
    Expression ="Sum(LeaseReport2.Other)"
    Alias ="TotalChildCost"
    Expression ="Sum(LeaseReport2.TotalChildCost)"
    Alias ="TotalCost"
    Expression ="Sum(LeaseReport2.TotalCost)"
End
Begin Groups
    Expression ="LeaseReport2.LeaseYear"
    GroupLevel =0
    Expression ="LeaseReport2.Lessee"
    GroupLevel =0
    Expression ="LeaseReport2.[Lease No]"
    GroupLevel =0
    Expression ="LeaseReport2.Make"
    GroupLevel =0
    Expression ="LeaseReport2.UnitType"
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
        dbText "Name" ="LeaseReport2.LeaseYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Refrig"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInstall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FeesTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Freight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Inspection"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OilRig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankBodies"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OtherCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCost"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport2.Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport2.UnitType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4875"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PTO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalChildCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Decals"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport2.[Lease No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport2.Lessee"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =48
    Top =35
    Right =2279
    Bottom =819
    Left =-1
    Top =-1
    Right =2199
    Bottom =301
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =87
        Top =9
        Right =273
        Bottom =267
        Top =0
        Name ="LeaseReport2"
        Name =""
    End
End
