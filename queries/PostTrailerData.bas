Operation =3
Name ="TrailerMilesHours"
Option =0
Begin InputTables
    Name ="TrailerDataImporter_Maines"
End
Begin OutputColumns
    Alias ="Expr1"
    Name ="UnitID"
    Expression ="TrailerDataImporter_Maines.UnitID"
    Alias ="Expr2"
    Name ="UnitNumber"
    Expression ="TrailerDataImporter_Maines.UnitUnitNum"
    Alias ="Expr3"
    Name ="LogActivityDate"
    Expression ="TrailerDataImporter_Maines.LogActivityDate"
    Alias ="Expr4"
    Name ="Miles"
    Expression ="TrailerDataImporter_Maines.Miles"
    Alias ="Expr5"
    Name ="MilesLTD"
    Expression ="TrailerDataImporter_Maines.MilesLTD"
    Alias ="Expr6"
    Name ="Hours"
    Expression ="TrailerDataImporter_Maines.Hours"
    Alias ="Expr7"
    Name ="HoursLTD"
    Expression ="TrailerDataImporter_Maines.HoursLTD"
    Alias ="Expr8"
    Name ="OBCDataStatus"
    Expression ="TrailerDataImporter_Maines.OBCDataStatus"
    Alias ="DataDesc"
    Name ="OBCDataDesc"
    Expression ="IIf([ObcDataDesc]=\"Reporting,Reporting\",\"Reporting\",[ObcDataDesc])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="ImportTrailerData.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImportTrailerData.NewLogActivityDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImportTrailerData.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImportTrailerData.MonthsBetween"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImportTrailerData.NewMiles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImportTrailerData.MILES"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImportTrailerData.NewHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ImportTrailerData.RFRHR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDataStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HoursLtdX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MIlesX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerDataImporter_Maines.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerDataImporter_Maines.MilesLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerDataImporter_Maines.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerDataImporter_Maines.MILES"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerDataImporter_Maines.hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerDataImporter_Maines.HoursLtd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerDataImporter_Maines.logactivitydate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerDataImporter_Maines.OBCDataStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TrailerDataImporter_Maines.ObcDataDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3105"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DataDesc"
        dbInteger "ColumnWidth" ="3105"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =52
    Top =62
    Right =1479
    Bottom =898
    Left =-1
    Top =-1
    Right =1395
    Bottom =358
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =633
        Top =89
        Right =909
        Bottom =360
        Top =0
        Name ="TrailerDataImporter_Maines"
        Name =""
    End
End
