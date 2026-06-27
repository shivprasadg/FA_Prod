Operation =1
Option =0
Having ="(((ChildParts.ChildType) Like \"*Dome*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="OBC_Dome"
    Expression ="ChildParts.ChildMake"
    Alias ="OBCDome"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="OBCDomeInv"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="OBCDomePaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="OBCDomeDraw"
    Expression ="ChildParts.DrawDownId"
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
        dbText "Name" ="OBC_Dome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDomeInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDomePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDomeDraw"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =53
    Top =49
    Right =1454
    Bottom =772
    Left =-1
    Top =-1
    Right =1369
    Bottom =283
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
