Operation =1
Option =0
Begin InputTables
    Name ="LeaseReport2"
End
Begin OutputColumns
    Expression ="LeaseReport2.Lessee"
    Expression ="LeaseReport2.[Lease No]"
    Alias ="UnitClass"
    Expression ="IIf([UnitType] Like \"*trailer*\",\"Trailer\",\"Tractor\")"
    Expression ="LeaseReport2.ModelYear"
    Expression ="LeaseReport2.Make"
    Expression ="LeaseReport2.UnitType"
    Alias ="OrderedYear"
    Expression ="CDbl(Nz([LeaseReport2].[OrderedYear],Null))"
    Alias ="OwnedYear"
    Expression ="CDbl(Nz([LeaseReport2].[OwnedYear],Null))"
    Alias ="SYDYear"
    Expression ="CDbl(Nz([LeaseReport2].[SYDYear],Null))"
    Expression ="LeaseReport2.Ordered"
    Alias ="OnOrderCount"
    Expression ="IIf([Ordered]=\"Ordered\",Count([AssetId]),0)"
    Alias ="OnOrderCost"
    Expression ="IIf([Ordered]=\"Ordered\",Sum(CCur(Nz([LeaseReport2].[UnitCost],0))),0)"
    Alias ="OnOrderOEC"
    Expression ="IIf([Ordered]=\"Ordered\",Sum(CCur(Nz([LeaseReport2].[TotalCost],0))),0)"
    Expression ="LeaseReport2.Owned"
    Alias ="OwnedCount"
    Expression ="IIf([Owned]=\"Owned\",Count([AssetId]),0)"
    Alias ="OwnedCost"
    Expression ="IIf([Owned]=\"Owned\",Sum(CCur(Nz([LeaseReport2].[UnitCost],0))),0)"
    Alias ="OwnedOEC"
    Expression ="IIf([Owned]=\"Owned\",Sum(CCur(Nz([LeaseReport2].[TotalCost],0))),0)"
    Expression ="LeaseReport2.Syndicated"
    Alias ="SydCount"
    Expression ="IIf([Syndicated]=\"Syndicated\",Count([AssetId]),0)"
    Alias ="SydCost"
    Expression ="IIf([Syndicated]=\"Syndicated\",Sum(CCur(Nz([LeaseReport2].[UnitCost],0))),0)"
    Alias ="SydOEC"
    Expression ="IIf([Syndicated]=\"Syndicated\",Sum(CCur(Nz([LeaseReport2].[TotalCost],0))),0)"
End
Begin OrderBy
    Expression ="IIf([UnitType] Like \"*trailer*\",\"Trailer\",\"Tractor\")"
    Flag =0
    Expression ="CDbl(Nz([LeaseReport2].[OrderedYear],Null))"
    Flag =1
    Expression ="CDbl(Nz([LeaseReport2].[OwnedYear],Null))"
    Flag =1
    Expression ="CDbl(Nz([LeaseReport2].[SYDYear],Null))"
    Flag =1
End
Begin Groups
    Expression ="LeaseReport2.Lessee"
    GroupLevel =0
    Expression ="LeaseReport2.[Lease No]"
    GroupLevel =0
    Expression ="IIf([UnitType] Like \"*trailer*\",\"Trailer\",\"Tractor\")"
    GroupLevel =0
    Expression ="LeaseReport2.ModelYear"
    GroupLevel =0
    Expression ="LeaseReport2.Make"
    GroupLevel =0
    Expression ="LeaseReport2.UnitType"
    GroupLevel =0
    Expression ="CDbl(Nz([LeaseReport2].[OrderedYear],Null))"
    GroupLevel =0
    Expression ="CDbl(Nz([LeaseReport2].[OwnedYear],Null))"
    GroupLevel =0
    Expression ="CDbl(Nz([LeaseReport2].[SYDYear],Null))"
    GroupLevel =0
    Expression ="LeaseReport2.Ordered"
    GroupLevel =0
    Expression ="LeaseReport2.Owned"
    GroupLevel =0
    Expression ="LeaseReport2.Syndicated"
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
        dbText "Name" ="SYDYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="LeaseReport2.UnitType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="OrderedYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="SydOEC"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="OnOrderCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="OnOrderOEC"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="LeaseReport2.Owned"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="OwnedYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="LeaseReport2.Ordered"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="SydCount"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport2.Syndicated"
        dbInteger "ColumnWidth" ="2760"
        dbInteger "ColumnOrder" ="16"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnOrderCount"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitClass"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport2.ModelYear"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OwnedCount"
        dbInteger "ColumnWidth" ="1785"
        dbInteger "ColumnOrder" ="13"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OwnedOEC"
        dbInteger "ColumnWidth" ="1725"
        dbInteger "ColumnOrder" ="14"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OwnedCost"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SydCost"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport2.Lessee"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport2.[Lease No]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="LeaseReport2.Make"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =121
    Top =100
    Right =2563
    Bottom =1051
    Left =-1
    Top =-1
    Right =2418
    Bottom =358
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =174
        Top =105
        Right =679
        Bottom =504
        Top =0
        Name ="LeaseReport2"
        Name =""
    End
End
