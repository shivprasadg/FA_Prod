Operation =1
Option =0
Where ="(((TrailerMilesHours.[UnitID]) In (SELECT [UnitID] FROM [TrailerMilesHours] As T"
    "mp GROUP BY [UnitID],[LogActivityDate] HAVING Count(*)>1  And [LogActivityDate] "
    "= [TrailerMilesHours].[LogActivityDate])))"
Begin InputTables
    Name ="TrailerMilesHours"
End
Begin OutputColumns
    Expression ="TrailerMilesHours.[UnitID]"
    Expression ="TrailerMilesHours.[LogActivityDate]"
    Expression ="TrailerMilesHours.[OBCDataStatus]"
    Expression ="TrailerMilesHours.[Miles]"
    Expression ="TrailerMilesHours.[Hours]"
    Expression ="TrailerMilesHours.ReadDate"
End
Begin OrderBy
    Expression ="TrailerMilesHours.[UnitID]"
    Flag =0
    Expression ="TrailerMilesHours.[LogActivityDate]"
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
Begin
    Begin
        dbText "Name" ="TrailerMilesHours.[OBCDataStatus]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerMilesHours.[UnitID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerMilesHours.ReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[UnitID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[LogActivityDate]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[OBCDataStatus]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Miles]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerMilesHours.[LogActivityDate]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Hours]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerMilesHours.[Miles]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerMilesHours.[Hours]"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =4
    Right =1609
    Bottom =990
    Left =-1
    Top =-1
    Right =1503
    Bottom =271
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="TrailerMilesHours"
        Name =""
    End
End
