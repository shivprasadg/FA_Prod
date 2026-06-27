Operation =1
Option =0
Having ="(((ChildParts.ChildInvoiceNum) Is Not Null) AND ((ChildParts.ChildType) Like \"*"
    "Oil*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="OilRig"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="OilRigMake"
    Expression ="ChildParts.ChildMake"
    Alias ="OilRigInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="OilRigPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Expression ="ChildParts.ChildType"
    Alias ="OilRigDraw"
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
    Expression ="ChildParts.ChildType"
    GroupLevel =0
    Expression ="ChildParts.DrawDownId"
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
        dbInteger "ColumnWidth" ="4245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OilRigMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OilRig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OilRigInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OilRigPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OilRigDraw"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =157
    Top =62
    Right =1445
    Bottom =785
    Left =-1
    Top =-1
    Right =1256
    Bottom =151
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =93
        Top =30
        Right =507
        Bottom =403
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
