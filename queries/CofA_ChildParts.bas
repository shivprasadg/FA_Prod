Operation =1
Option =0
Where ="(((ChildParts.ChildEXCofa)=No))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Expression ="ChildParts.ChildYear"
    Expression ="ChildParts.ChildType"
    Expression ="ChildParts.ChildMake"
    Expression ="ChildParts.ChildClass"
    Expression ="ChildParts.ChildVIN"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="Cost"
    Expression ="Nz([childFinalCost],0)"
    Alias ="DatePaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="EngSerialNum"
    Expression ="ChildParts.REFENG"
    Alias ="RefrigDrawDownId"
    Expression ="ChildParts.DrawDownId"
    Expression ="ChildParts.ChildEXCofa"
    Expression ="ChildParts.ChildPartID"
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
        dbText "Name" ="AssetId"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigDrawDownId"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4050"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ChildParts.ChildVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildMake"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildModel"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildInvoiceNum"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DatePaid"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EngSerialNum"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildEXCofa"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildPartID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =177
    Top =82
    Right =1465
    Bottom =805
    Left =-1
    Top =-1
    Right =1264
    Bottom =185
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =93
        Top =30
        Right =328
        Bottom =403
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
