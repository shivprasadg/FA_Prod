Operation =1
Option =0
Begin InputTables
    Name ="TrailerMilesHours"
    Name ="TrailerLastReportDate"
End
Begin OutputColumns
    Expression ="TrailerMilesHours.UnitID"
    Expression ="TrailerMilesHours.LogActivityDate"
    Expression ="TrailerMilesHours.MilesLTD"
    Expression ="TrailerMilesHours.HoursLTD"
End
Begin Joins
    LeftTable ="TrailerMilesHours"
    RightTable ="TrailerLastReportDate"
    Expression ="TrailerMilesHours.UnitID = TrailerLastReportDate.UnitID"
    Flag =1
    LeftTable ="TrailerMilesHours"
    RightTable ="TrailerLastReportDate"
    Expression ="TrailerMilesHours.LogActivityDate = TrailerLastReportDate.MaxOfLogActivityDate"
    Flag =1
End
Begin Groups
    Expression ="TrailerMilesHours.UnitID"
    GroupLevel =0
    Expression ="TrailerMilesHours.LogActivityDate"
    GroupLevel =0
    Expression ="TrailerMilesHours.MilesLTD"
    GroupLevel =0
    Expression ="TrailerMilesHours.HoursLTD"
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
        dbText "Name" ="TrailerMilesHours.LogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerMilesHours.HoursLTD"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =96
    Top =7
    Right =1320
    Bottom =988
    Left =-1
    Top =-1
    Right =1192
    Bottom =629
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =333
        Top =141
        Right =477
        Bottom =502
        Top =0
        Name ="TrailerMilesHours"
        Name =""
    End
    Begin
        Left =581
        Top =166
        Right =725
        Bottom =310
        Top =0
        Name ="TrailerLastReportDate"
        Name =""
    End
End
