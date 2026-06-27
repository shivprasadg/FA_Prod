Operation =1
Option =0
Begin InputTables
    Name ="SELECT UnitID, UnitVIN, Right([UnitVin],8) AS Veight FROM Units WHERE UnitStatus"
        "=\"EM\""
    Alias ="U"
    Name ="MoveOrderImport"
End
Begin OutputColumns
    Expression ="U.UnitID"
    Expression ="MoveOrderImport.[Order#]"
    Expression ="MoveOrderImport.EntryDate"
    Expression ="MoveOrderImport.LoadDate"
    Expression ="MoveOrderImport.ETADate"
    Expression ="MoveOrderImport.Status"
    Expression ="MoveOrderImport.Miles"
End
Begin Joins
    LeftTable ="U"
    RightTable ="MoveOrderImport"
    Expression ="U.Veight = MoveOrderImport.[Serial#]"
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
        dbText "Name" ="MoveOrderImport.EntryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MoveOrderImport.[Order#]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="U.UnitID"
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
End
Begin
    State =0
    Left =0
    Top =40
    Right =1472
    Bottom =1247
    Left =-1
    Top =-1
    Right =1448
    Bottom =859
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =440
        Top =84
        Right =584
        Bottom =228
        Top =0
        Name ="U"
        Name =""
    End
    Begin
        Left =633
        Top =231
        Right =777
        Bottom =542
        Top =0
        Name ="MoveOrderImport"
        Name =""
    End
End
