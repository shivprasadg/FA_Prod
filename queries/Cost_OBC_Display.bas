Operation =1
Option =0
Having ="(((ChildParts.ChildType) Like \"*Display*\" And (ChildParts.ChildType) Not Like "
    "\"*Installation*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="OBC_Display"
    Expression ="ChildParts.ChildMake"
    Alias ="OBCDisplay"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="OBCDisplayInv"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="OBCDisplayPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="OBCDisplayDraw"
    Expression ="ChildParts.DrawDownId"
    Expression ="ChildParts.ChildType"
End
Begin Groups
    Expression ="ChildParts.UnitREF"
    GroupLevel =0
    Expression ="ChildParts.ChildMake"
    GroupLevel =0
    Expression ="ChildParts.ChildInvoiceNum"
    GroupLevel =0
    Expression ="ChildParts.ChildVendorPytDate"
    GroupLevel =0
    Expression ="ChildParts.DrawDownId"
    GroupLevel =0
    Expression ="ChildParts.ChildType"
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
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC_Adapt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCAdapt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCAdaptInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCAdaptPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCAdaptDraw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDisplay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDisplayPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDisplayInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC_Display"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDisplayDraw"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =147
    Top =131
    Right =1548
    Bottom =854
    Left =-1
    Top =-1
    Right =1369
    Bottom =181
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =178
        Top =43
        Right =384
        Bottom =389
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
