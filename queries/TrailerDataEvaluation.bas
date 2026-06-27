Operation =1
Option =0
Where ="(((MainesTrailersDataSource.LogActivityDate)>[TrailerMilesHours].[LogActivityDat"
    "e]))"
Begin InputTables
    Name ="MainesTrailersDataSource"
    Name ="TrailerMilesHours"
End
Begin OutputColumns
    Expression ="TrailerMilesHours.UnitID"
    Alias ="ReadDate"
    Expression ="Max(MainesTrailersDataSource.ReadDateMiles)"
    Alias ="LastDate"
    Expression ="Max(TrailerMilesHours.LogActivityDate)"
    Alias ="NewDate"
    Expression ="Max(MainesTrailersDataSource.LogActivityDate)"
    Alias ="PrevMilesLTD"
    Expression ="Max(TrailerMilesHours.MilesLTD)"
    Alias ="CurrMilesLTD"
    Expression ="Max(MainesTrailersDataSource.Miles)"
    Alias ="MilesDriven"
    Expression ="[CurrMilesLTD]-[PrevMilesLTD]"
    Alias ="MonthDiff"
    Expression ="DateDiff(\"m\",[LastDate],[NewDate])"
    Alias ="MilesLimit"
    Expression ="([MonthDiff]*14000)"
    Alias ="MilesDataStatus"
    Expression ="Switch([MilesDriven]=0,0,[MilesDriven]>[MilesLimit],2,[MilesDriven]<0,3,[MilesDr"
        "iven]>0,1)"
    Alias ="MilesDataStatusDesc"
    Expression ="Switch([MilesDataStatus]=2,\"Excess Miles: \" & ([MilesDriven]-[MilesLimit]),[Mi"
        "lesDataStatus]=3,\"Neg Miles\",[MilesDataStatus]=0,\"No Miles\",[MilesDataStatus"
        "]=1,\"Reported Miles\")"
    Alias ="MilesResult"
    Expression ="IIf([MilesDriven]<=0,0,[MilesDriven])"
    Alias ="PrevHoursLTD"
    Expression ="Max(TrailerMilesHours.HoursLTD)"
    Alias ="CurrHoursLTD"
    Expression ="Max(MainesTrailersDataSource.Hours)"
    Alias ="HoursUsed"
    Expression ="[CurrHoursLTD]-[PrevHoursLTD]"
    Alias ="HoursLimit"
    Expression ="[MonthDiff]*(24*30)"
    Alias ="HoursDataStatus"
    Expression ="Switch([HoursUsed]>[HoursLimit],2,[HoursUsed]<0,3,[HoursUsed]=0,0,[HoursUsed]>0,"
        "1)"
    Alias ="HoursDataStatusDesc"
    Expression ="Switch([HoursDataStatus]=2,\"Excess Hours: \" & [HoursUsed]-[HoursLimit],[HoursD"
        "ataStatus]=3,\"Neg Hours\",[HoursDataStatus]=0,\"No Hours\",[HoursDataStatus]=1,"
        "\"Reported Hours\")"
    Alias ="ObcDataDesc"
    Expression ="[MilesDataStatusDesc] & \", \" & [HoursDataStatusDesc]"
    Alias ="DataDesc"
    Expression ="IIf([ObcDataDesc]=\"Reported Miles, Reported Hours\",\"Reporting\",[ObcDataDesc]"
        ")"
    Alias ="OBCDataStatus"
    Expression ="IIf([MilesDataStatus]+[HoursDataStatus] Between 1 And 2,1,[MilesDataStatus]+[Hou"
        "rsDataStatus])"
End
Begin Joins
    LeftTable ="MainesTrailersDataSource"
    RightTable ="TrailerMilesHours"
    Expression ="MainesTrailersDataSource.UnitID = TrailerMilesHours.UnitID"
    Flag =2
End
Begin Groups
    Expression ="TrailerMilesHours.UnitID"
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
        dbText "Name" ="LastDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MonthDiff"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[MilesDriven]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MilesDataStatus"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MilesLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MilesDriven"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MilesDataStatusDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2355"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="MilesResult"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrHoursLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevMilesLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HoursUsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrevHoursLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CurrMilesLTD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HoursDataStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HoursDataStatusDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HoursLimit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ObcDataDesc"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDataStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DataDesc"
        dbInteger "ColumnWidth" ="3570"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ReadDate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =174
    Top =104
    Right =1695
    Bottom =942
    Left =-1
    Top =-1
    Right =1489
    Bottom =416
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =42
        Top =4
        Right =201
        Bottom =269
        Top =0
        Name ="MainesTrailersDataSource"
        Name =""
    End
    Begin
        Left =477
        Top =123
        Right =621
        Bottom =399
        Top =0
        Name ="TrailerMilesHours"
        Name =""
    End
End
