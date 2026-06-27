dbMemo "SQL" ="SELECT ChildParts.UnitREF, ChildParts.ChildVIN, ChildParts.ChildYear, ChildParts"
    ".ChildMake, ChildParts.ChildClass, ChildParts.ChildType, ChildParts.ChildEXCofa,"
    " IIf(IsNull([ChildVIN]),Null,Trim([childyear]) & \"   \" & Trim([childmake]) & I"
    "If([ChildClass]=\".\",Null,\"   \" & Trim([ChildClass])) & \"   \" & Trim([child"
    "type]) & \"    \" & Trim([ChildVIN])) AS ChildDesc\015\012FROM ChildParts\015\012"
    "WHERE (((ChildParts.ChildEXCofa)=No))\015\012ORDER BY ChildParts.UnitREF, ChildP"
    "arts.ChildVIN;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="ChildParts.UnitREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildEXCofa"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildDesc"
        dbInteger "ColumnWidth" ="7005"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
