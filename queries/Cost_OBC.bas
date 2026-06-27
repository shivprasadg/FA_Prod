Operation =1
Option =0
Having ="(((ChildParts.ChildType) Like \"*OBC*\" And (ChildParts.ChildType) Not Like \"*I"
    "nstall*\" And (ChildParts.ChildType) Not Like \"*Display*\" And (ChildParts.Chil"
    "dType) Not Like \"*Part*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="OBC_Make"
    Expression ="ChildParts.ChildMake"
    Alias ="OBC"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="OBCInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="OBCPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="OBCDraw"
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
dbMemo "Filter" ="([Cost_OBC].[AssetId]=285305)"
Begin
    Begin
        dbText "Name" ="OBC_Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDraw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbInteger "ColumnWidth" ="3150"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =156
    Top =185
    Right =1557
    Bottom =908
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
