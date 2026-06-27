Operation =1
Option =0
Having ="(((ChildParts.ChildType) Like \"*OBC*\" And (ChildParts.ChildType) Like \"*Insta"
    "ll*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="OBCInstall"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="OBCInsInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="OBCInstaller"
    Expression ="Nz([ChildMake],\"Other\")"
    Alias ="OBCInstallPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="OBCInstallDraw"
    Expression ="ChildParts.DrawDownId"
    Expression ="ChildParts.ChildType"
End
Begin Groups
    Expression ="ChildParts.UnitREF"
    GroupLevel =0
    Expression ="ChildParts.ChildInvoiceNum"
    GroupLevel =0
    Expression ="Nz([ChildMake],\"Other\")"
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
        dbText "Name" ="OBCInstaller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInsInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInstall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInstallPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInstallDraw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =126
    Top =125
    Right =1528
    Bottom =848
    Left =-1
    Top =-1
    Right =1370
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
