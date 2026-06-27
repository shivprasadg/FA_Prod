Operation =1
Option =0
Begin InputTables
    Name ="vw_Sch_Grp_Units_ChildParts"
    Alias ="DCP"
End
Begin OutputColumns
    Expression ="DCP.ChildPartID"
    Expression ="DCP.AssetID"
    Expression ="DCP.VIN"
    Expression ="DCP.Unitnum"
    Alias ="Invoice"
    Expression ="DCP.ChildInvoiceNum"
    Expression ="DCP.ChildType"
    Expression ="DCP.ChildCost"
    Alias ="GroupID"
    Expression ="DCP.SGrpID"
    Expression ="DCP.ChildMake"
    Expression ="DCP.ChildVendorID"
    Expression ="DCP.Schedule"
    Expression ="DCP.FAID"
    Expression ="DCP.SchID"
    Expression ="DCP.ChildDrawDownId"
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
        dbText "Name" ="Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.ChildVendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.Unitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.ChildDrawDownId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DCP.VIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.ChildCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.ChildPartID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DCP.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =425
    Top =112
    Right =1766
    Bottom =927
    Left =-1
    Top =-1
    Right =1317
    Bottom =314
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="DCP"
        Name =""
    End
End
