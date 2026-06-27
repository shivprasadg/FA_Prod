Operation =3
Name ="ChildParts"
Option =0
Where ="(((Units.GroupREF)=[forms]![add additional child]![selgrp]))"
Begin InputTables
    Name ="Units"
End
Begin OutputColumns
    Name ="UnitREF"
    Expression ="Units.UnitID"
    Alias ="YR"
    Name ="ChildYear"
    Expression ="[forms]![add additional Child]![AddCYR]"
    Alias ="Make"
    Name ="ChildMake"
    Expression ="[forms]![add additional Child]![AddCMake]"
    Alias ="Model"
    Name ="ChildClass"
    Expression ="[forms]![add additional Child]![AddcModel]"
    Alias ="Type"
    Name ="ChildType"
    Expression ="[forms]![add additional Child]![AddCType]"
    Alias ="Cost"
    Name ="childFinalCost"
    Expression ="[forms]![add additional Child]![AddCCost]"
    Alias ="Vendor"
    Name ="ChildVendorREF"
    Expression ="[forms]![add additional Child]![AddCvendor]"
    Alias ="Status"
    Name ="ChildStatus"
    Expression ="[forms]![add additional Child]![AddCstatus]"
    Alias ="OnCofA"
    Name ="ChildEXCofa"
    Expression ="[forms]![add additional Child]![AccCCofa]"
    Alias ="CHAdd"
    Name ="PartAddedDate"
    Expression ="Date()"
End
Begin OrderBy
    Expression ="Units.UnitID"
    Flag =0
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="YR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Model"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OnCofA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CHAdd"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =343
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =424
        Bottom =214
        Top =0
        Name ="Units"
        Name =""
    End
End
