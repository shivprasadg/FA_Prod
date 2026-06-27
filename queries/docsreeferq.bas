Operation =1
Option =0
Where ="(((ChildParts.ChildType)=\"refrig unit\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Expression ="ChildParts.UnitREF"
    Alias ="Serial Number"
    Expression ="ChildParts.ChildVIN"
    Expression ="ChildParts.ChildMake"
    Expression ="ChildParts.ChildClass"
    Expression ="ChildParts.ChildType"
    Alias ="RefInvoiceNum"
    Expression ="ChildParts.ChildInvoiceNum"
End
Begin OrderBy
    Expression ="ChildParts.UnitREF"
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
        dbText "Name" ="ChildParts.UnitREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Serial Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefInvoiceNum"
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
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =294
        Bottom =238
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
