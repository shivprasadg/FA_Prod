Operation =1
Option =0
Having ="(((ChildParts.ChildInvoiceNum) Is Not Null) AND ((ChildParts.ChildType)<>\"Tank "
    "Body\" And (ChildParts.ChildType) Like \"*Tank*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="TankOther"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="TankOthMake"
    Expression ="ChildParts.ChildMake"
    Alias ="TankOthInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="TankOthPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Expression ="ChildParts.ChildType"
    Alias ="TankOthDraw"
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
        dbText "Name" ="TankOthDraw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankOthPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankOthInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankOthMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankOther"
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
    Bottom =185
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
