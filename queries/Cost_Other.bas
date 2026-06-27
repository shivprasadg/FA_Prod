Operation =1
Option =0
Having ="(((ChildParts.ChildInvoiceNum) Is Not Null) AND ((ChildParts.ChildType) Not Like"
    " \"*Refrig*\" And (ChildParts.ChildType) Not Like \"*Decal*\" And (ChildParts.Ch"
    "ildType) Not Like \"*Freight*\" And (ChildParts.ChildType) Not Like \"*OBC*\" An"
    "d (ChildParts.ChildType) Not Like \"*XATA*\" And (ChildParts.ChildType) Not Like"
    " \"*Warranty*\" And (ChildParts.ChildType) Not Like \"*Body*\" And (ChildParts.C"
    "hildType) Not Like \"*Tank*\" And (ChildParts.ChildType) Not Like \"*Tax*\" And "
    "(ChildParts.ChildType) Not Like \"*Inspection*\" And (ChildParts.ChildType) Not "
    "Like \"*PTO*\" And (ChildParts.ChildType) Not Like \"*Oil*\" And (ChildParts.Chi"
    "ldType) Not Like \"*Cable*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Alias ="AssetId"
    Expression ="ChildParts.UnitREF"
    Alias ="Other"
    Expression ="Sum(CCur(Nz([childFinalCost],0)))"
    Alias ="OtherMake"
    Expression ="[ChildMake] & \" - \" & [ChildType]"
    Alias ="OtherInvoice"
    Expression ="ChildParts.ChildInvoiceNum"
    Alias ="OtherPaid"
    Expression ="ChildParts.ChildVendorPytDate"
    Alias ="OtherDraw"
    Expression ="ChildParts.DrawDownId"
    Expression ="ChildParts.ChildType"
End
Begin Groups
    Expression ="ChildParts.UnitREF"
    GroupLevel =0
    Expression ="[ChildMake] & \" - \" & [ChildType]"
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
        dbText "Name" ="Other"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OtherMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3630"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="OtherInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OtherPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OtherDraw"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =214
    Top =57
    Right =1502
    Bottom =780
    Left =-1
    Top =-1
    Right =1256
    Bottom =316
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
