Operation =1
Option =0
Having ="(((ChildParts.ChildType) Like \"*Connection*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="Connect_Make"
    Expression ="ChildParts.ChildMake"
    Alias ="Connect"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="ConnectInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="ConnectPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="ConnectDraw"
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
        dbText "Name" ="ConnectPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ConnectInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Connect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Connect_Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ConnectDraw"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =127
    Top =125
    Right =1528
    Bottom =848
    Left =-1
    Top =-1
    Right =1369
    Bottom =300
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
