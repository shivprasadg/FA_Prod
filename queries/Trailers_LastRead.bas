Operation =1
Option =0
Having ="(((Max(TrailerMilesHours.LogActivityDate))<#1/1/2016#))"
Begin InputTables
    Name ="TrailerMilesHours"
End
Begin OutputColumns
    Expression ="TrailerMilesHours.UnitID"
    Alias ="LastRead"
    Expression ="Max(TrailerMilesHours.LogActivityDate)"
    Expression ="TrailerMilesHours.MilesLTD"
End
Begin OrderBy
    Expression ="Max(TrailerMilesHours.LogActivityDate)"
    Flag =1
End
Begin Groups
    Expression ="TrailerMilesHours.UnitID"
    GroupLevel =0
    Expression ="TrailerMilesHours.MilesLTD"
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
        dbText "Name" ="TrailerMilesHours.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerMilesHours.MilesLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastRead"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1350
    Bottom =974
    Left =-1
    Top =-1
    Right =1318
    Bottom =633
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =110
        Top =71
        Right =593
        Bottom =459
        Top =0
        Name ="TrailerMilesHours"
        Name =""
    End
End
