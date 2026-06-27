Operation =1
Option =0
Having ="(((ChildParts.ChildType) Like \"*OBC Part*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="OBC_Part"
    Expression ="ChildParts.ChildMake"
    Alias ="OBCPart"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="OBCPartInv"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="OBCPartPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="OBCPartDraw"
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
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPartDraw"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPartInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC_Part"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPartPaid"
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
    Bottom =130
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
