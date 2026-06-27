Operation =1
Option =0
Begin InputTables
    Name ="StagingVendorTracking"
End
Begin OutputColumns
    Expression ="StagingVendorTracking.Id"
    Expression ="StagingVendorTracking.VendorId"
    Expression ="StagingVendorTracking.ClientId"
    Expression ="StagingVendorTracking.PO"
    Expression ="StagingVendorTracking.FAID"
    Expression ="StagingVendorTracking.UnitId"
    Expression ="StagingVendorTracking.UnitBuildDate"
    Expression ="StagingVendorTracking.UnitEstDelDate"
    Expression ="StagingVendorTracking.ActualDelivery"
    Expression ="StagingVendorTracking.UnitType"
    Expression ="StagingVendorTracking.UnitYr"
    Expression ="StagingVendorTracking.UnitMake"
    Expression ="StagingVendorTracking.UnitModel"
    Expression ="StagingVendorTracking.FinalCost"
    Expression ="StagingVendorTracking.Axle"
    Expression ="StagingVendorTracking.VIN"
    Expression ="StagingVendorTracking.EngineSer"
    Expression ="StagingVendorTracking.TransSer"
    Expression ="StagingVendorTracking.FrontAxleSer"
    Expression ="StagingVendorTracking.RAxleFWDSer"
    Expression ="StagingVendorTracking.RAxleREARSer"
    Expression ="StagingVendorTracking.REFENG"
    Expression ="StagingVendorTracking.COMPRESSS"
    Expression ="StagingVendorTracking.MICROPROC"
    Expression ="StagingVendorTracking.ELECDISP"
    Expression ="StagingVendorTracking.DeliverTo"
    Expression ="StagingVendorTracking.UnitNotes"
    Expression ="StagingVendorTracking.IsPilot"
    Expression ="StagingVendorTracking.DeliveryLocation"
    Expression ="StagingVendorTracking.ChildType"
    Expression ="StagingVendorTracking.ChildMake"
    Expression ="StagingVendorTracking.ChildYear"
    Expression ="StagingVendorTracking.Freight"
    Expression ="StagingVendorTracking.ChildPartID"
    Alias ="LatestUpdate"
    Expression ="Max(StagingVendorTracking.ImportDate)"
End
Begin Groups
    Expression ="StagingVendorTracking.Id"
    GroupLevel =0
    Expression ="StagingVendorTracking.VendorId"
    GroupLevel =0
    Expression ="StagingVendorTracking.ClientId"
    GroupLevel =0
    Expression ="StagingVendorTracking.PO"
    GroupLevel =0
    Expression ="StagingVendorTracking.FAID"
    GroupLevel =0
    Expression ="StagingVendorTracking.UnitId"
    GroupLevel =0
    Expression ="StagingVendorTracking.UnitBuildDate"
    GroupLevel =0
    Expression ="StagingVendorTracking.UnitEstDelDate"
    GroupLevel =0
    Expression ="StagingVendorTracking.ActualDelivery"
    GroupLevel =0
    Expression ="StagingVendorTracking.UnitType"
    GroupLevel =0
    Expression ="StagingVendorTracking.UnitYr"
    GroupLevel =0
    Expression ="StagingVendorTracking.UnitMake"
    GroupLevel =0
    Expression ="StagingVendorTracking.UnitModel"
    GroupLevel =0
    Expression ="StagingVendorTracking.FinalCost"
    GroupLevel =0
    Expression ="StagingVendorTracking.Axle"
    GroupLevel =0
    Expression ="StagingVendorTracking.VIN"
    GroupLevel =0
    Expression ="StagingVendorTracking.EngineSer"
    GroupLevel =0
    Expression ="StagingVendorTracking.TransSer"
    GroupLevel =0
    Expression ="StagingVendorTracking.FrontAxleSer"
    GroupLevel =0
    Expression ="StagingVendorTracking.RAxleFWDSer"
    GroupLevel =0
    Expression ="StagingVendorTracking.RAxleREARSer"
    GroupLevel =0
    Expression ="StagingVendorTracking.REFENG"
    GroupLevel =0
    Expression ="StagingVendorTracking.COMPRESSS"
    GroupLevel =0
    Expression ="StagingVendorTracking.MICROPROC"
    GroupLevel =0
    Expression ="StagingVendorTracking.ELECDISP"
    GroupLevel =0
    Expression ="StagingVendorTracking.DeliverTo"
    GroupLevel =0
    Expression ="StagingVendorTracking.UnitNotes"
    GroupLevel =0
    Expression ="StagingVendorTracking.IsPilot"
    GroupLevel =0
    Expression ="StagingVendorTracking.DeliveryLocation"
    GroupLevel =0
    Expression ="StagingVendorTracking.ChildType"
    GroupLevel =0
    Expression ="StagingVendorTracking.ChildMake"
    GroupLevel =0
    Expression ="StagingVendorTracking.ChildYear"
    GroupLevel =0
    Expression ="StagingVendorTracking.Freight"
    GroupLevel =0
    Expression ="StagingVendorTracking.ChildPartID"
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
        dbText "Name" ="LatestUpdate"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.Id"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.VendorId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.ClientId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.PO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.UnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.UnitEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.ActualDelivery"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.UnitYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.FinalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.EngineSer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.TransSer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.FrontAxleSer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.RAxleFWDSer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.RAxleREARSer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.REFENG"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.COMPRESSS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.MICROPROC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.ELECDISP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.DeliverTo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.UnitNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.IsPilot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.DeliveryLocation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.ChildYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.Freight"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.ChildPartID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StagingVendorTracking.UnitId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1612
    Bottom =855
    Left =-1
    Top =-1
    Right =1542
    Bottom =435
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =103
        Top =24
        Right =373
        Bottom =403
        Top =0
        Name ="StagingVendorTracking"
        Name =""
    End
End
