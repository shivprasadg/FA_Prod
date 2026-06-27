Operation =1
Option =0
Where ="(((ChildParts.ChildType) Like \"*Inspection*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="Inspection"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="InspCo"
    Expression ="ChildParts.ChildMake"
    Alias ="InspInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="InspPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="InspDraw"
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
        dbText "Name" ="Inspection"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InspInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InspCo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InspPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InspDraw"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =187
    Top =132
    Right =1475
    Bottom =855
    Left =-1
    Top =-1
    Right =1256
    Bottom =270
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
