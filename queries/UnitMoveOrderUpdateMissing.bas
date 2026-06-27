Operation =1
Option =0
Where ="(((U.UnitID) Is Null))"
Begin InputTables
    Name ="SELECT UnitID, UnitVIN, Right([UnitVin],8) AS Veight FROM Units WHERE UnitStatus"
        " Like \"E*\""
    Alias ="U"
    Name ="MoveOrderImport"
End
Begin OutputColumns
    Expression ="MoveOrderImport.[Order#]"
    Expression ="MoveOrderImport.[Serial#]"
    Expression ="MoveOrderImport.[Driver#]"
    Expression ="MoveOrderImport.EntryDate"
    Expression ="MoveOrderImport.LoadDate"
    Expression ="MoveOrderImport.ETADate"
    Expression ="MoveOrderImport.PreAssignedDate"
    Expression ="MoveOrderImport.[Rev Terminal]"
    Expression ="MoveOrderImport.Status"
    Expression ="MoveOrderImport.[Origin Dest]"
    Expression ="MoveOrderImport.[Shipper Name]"
    Expression ="MoveOrderImport.[Consignee Name]"
    Expression ="MoveOrderImport.Miles"
    Expression ="MoveOrderImport.Commodity"
    Expression ="MoveOrderImport.ELD"
    Expression ="MoveOrderImport.TOWABLE"
    Expression ="MoveOrderImport.InStorage"
End
Begin Joins
    LeftTable ="U"
    RightTable ="MoveOrderImport"
    Expression ="U.Veight = MoveOrderImport.[Serial#]"
    Flag =3
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
        dbText "Name" ="MoveOrderImport.EntryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.[Order#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.LoadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.ETADate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.Miles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.ELD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.InStorage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.[Serial#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.[Driver#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.PreAssignedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.[Rev Terminal]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.[Origin Dest]"
        dbInteger "ColumnWidth" ="3660"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.[Shipper Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.[Consignee Name]"
        dbInteger "ColumnWidth" ="3090"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.Commodity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.TOWABLE"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =210
    Top =141
    Right =1845
    Bottom =1348
    Left =-1
    Top =-1
    Right =1611
    Bottom =808
    Left =120
    Top =0
    ColumnsShown =539
    Begin
        Left =320
        Top =84
        Right =464
        Bottom =228
        Top =0
        Name ="U"
        Name =""
    End
    Begin
        Left =513
        Top =231
        Right =657
        Bottom =542
        Top =0
        Name ="MoveOrderImport"
        Name =""
    End
End
