Operation =1
Option =0
Where ="(((Mileage.OBCDataStatus)=1 Or (Mileage.OBCDataStatus)=0))"
Begin InputTables
    Name ="Mileage"
    Name ="Units"
End
Begin OutputColumns
    Expression ="Units.UnitID"
    Alias ="ModelYear"
    Expression ="Units.UnitYr"
    Alias ="ReadDate"
    Expression ="Max(Mileage.LogActivityDate)"
    Alias ="Mileage"
    Expression ="Max(Nz([OdometerEnd],0))"
    Alias ="Hours"
    Expression ="Max(Nz([HoursLTD],0))"
End
Begin Joins
    LeftTable ="Mileage"
    RightTable ="Units"
    Expression ="Mileage.UnitId = Units.UnitID"
    Flag =3
End
Begin Groups
    Expression ="Units.UnitID"
    GroupLevel =0
    Expression ="Units.UnitYr"
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
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ModelYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReadDate"
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
    Bottom =347
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =301
        Top =1
        Right =683
        Bottom =423
        Top =0
        Name ="Mileage"
        Name =""
    End
    Begin
        Left =60
        Top =-8
        Right =204
        Bottom =344
        Top =0
        Name ="Units"
        Name =""
    End
End
