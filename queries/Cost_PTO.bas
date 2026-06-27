Operation =1
Option =0
Having ="(((ChildParts.ChildInvoiceNum) Is Not Null) AND ((ChildParts.ChildType) Like \"*"
    "PTO*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="PTO"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="PTOMake"
    Expression ="ChildParts.ChildMake"
    Alias ="PTOInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Expression ="ChildParts.ChildType"
    Alias ="PTOPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="PTODraw"
    Expression ="ChildParts.DrawDownId"
End
Begin Groups
    Expression ="ChildParts.UnitREF"
    GroupLevel =0
    Expression ="ChildParts.ChildMake"
    GroupLevel =0
    Expression ="ChildParts.ChildInvoiceNum"
    GroupLevel =0
    Expression ="ChildParts.ChildType"
    GroupLevel =0
    Expression ="ChildParts.ChildVendorPytDate"
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
        dbText "Name" ="PTOMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PTOInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PTO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PTOPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PTODraw"
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
    Bottom =236
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
