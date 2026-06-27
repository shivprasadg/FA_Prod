Operation =1
Option =0
Where ="(((ChildParts.ChildType) Not Like \"decal*\" And (ChildParts.ChildType) Not Like"
    " \"*warranty\" And (ChildParts.ChildType) Not Like \"*Installation*\" And (Child"
    "Parts.ChildType)<>\"freight\" And (ChildParts.ChildType) Not Like \"*modificatio"
    "n*\"))"
Begin InputTables
    Name ="ChildParts"
End
Begin OutputColumns
    Expression ="ChildParts.UnitREF"
    Expression ="ChildParts.ChildYear"
    Alias ="CPSerial No"
    Expression ="ChildParts.ChildVIN"
    Expression ="ChildParts.ChildMake"
    Expression ="ChildParts.ChildClass"
    Expression ="ChildParts.ChildType"
    Expression ="ChildParts.REFENG"
    Expression ="ChildParts.COMPRESSS"
    Expression ="ChildParts.MICROPROC"
    Expression ="ChildParts.ELECDISP"
    Expression ="ChildParts.[ARB ID]"
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
        dbText "Name" ="ChildParts.UnitREF"
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
        dbText "Name" ="ChildParts.REFENG"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.COMPRESSS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.MICROPROC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ELECDISP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.[ARB ID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CPSerial No"
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
    Bottom =369
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =492
        Bottom =156
        Top =0
        Name ="ChildParts"
        Name =""
    End
End
