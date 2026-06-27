Operation =1
Option =0
Where ="(((BillAdjT.BillAdjBillref)=[forms]![BillingSubForm]![billid]))"
Begin InputTables
    Name ="BillAdjT"
End
Begin OutputColumns
    Expression ="BillAdjT.BillAdjID"
    Expression ="BillAdjT.BillAdjBillref"
    Expression ="BillAdjT.BillAdj"
    Expression ="BillAdjT.BillAdjtype"
    Expression ="BillAdjT.BillAdjVIN"
    Expression ="BillAdjT.BillAdjUnitNo"
    Expression ="BillAdjT.BillAdjNote"
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
        dbText "Name" ="BillAdjT.BillAdjBillref"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2385"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="BillAdjT.BillAdjID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillAdjT.BillAdj"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillAdjT.BillAdjtype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillAdjT.BillAdjVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillAdjT.BillAdjUnitNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillAdjT.BillAdjNote"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =122
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="BillAdjT"
        Name =""
    End
End
