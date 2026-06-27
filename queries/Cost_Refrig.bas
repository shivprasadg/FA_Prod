Operation =1
Option =0
Where ="(((ChildParts.ChildType) Like \"*Refrig*\") AND ((ChildParts.ExcludeCostDocsRepo"
    "rts)=False))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="RefrigMake"
    Expression ="ChildParts.ChildMake"
    Alias ="RefrigModel"
    Expression ="ChildParts.ChildClass"
    Alias ="RefrigInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="Refrig"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="RefrigPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="RefrigSerialNum"
    Expression ="ChildParts.ChildVIN"
    Alias ="RefrigDrawDownId"
    Expression ="ChildParts.DrawDownId"
End
Begin Groups
    Expression ="ChildParts.UnitREF"
    GroupLevel =0
    Expression ="ChildParts.ChildMake"
    GroupLevel =0
    Expression ="ChildParts.ChildClass"
    GroupLevel =0
    Expression ="ChildParts.ChildInvoiceNum"
    GroupLevel =0
    Expression ="ChildParts.ChildVendorPytDate"
    GroupLevel =0
    Expression ="ChildParts.ChildVIN"
    GroupLevel =0
    Expression ="ChildParts.DrawDownId"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
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
        dbText "Name" ="Refrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigSerialNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigDrawDownId"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-335
    Top =140
    Right =953
    Bottom =863
    Left =-1
    Top =-1
    Right =1256
    Bottom =287
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
