Operation =1
Option =0
Where ="(((vw_sixkeys.MLOrig)=\"FA\"))"
Having ="(((Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\""
    ",[Type],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler"
    "\" Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" A"
    "nd [SubType]=\"TBD\",\"Tractor TBD\")) Is Not Null))"
Begin InputTables
    Name ="vw_sixkeys"
End
Begin OutputColumns
    Alias ="EOLYear"
    Expression ="DatePart(\"yyyy\",CDate(Nz([OffLeaseDate],Nz([ProjEOL],\"12/31/2050\"))))"
    Alias ="UnitType"
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
    Alias ="EOLCount"
    Expression ="Count(vw_sixkeys.AssetID)"
End
Begin OrderBy
    Expression ="DatePart(\"yyyy\",CDate(Nz([OffLeaseDate],Nz([ProjEOL],\"12/31/2050\"))))"
    Flag =0
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
    Flag =0
End
Begin Groups
    Expression ="DatePart(\"yyyy\",CDate(Nz([OffLeaseDate],Nz([ProjEOL],\"12/31/2050\"))))"
    GroupLevel =0
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
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
dbBoolean "TotalsRow" ="-1"
Begin
    Begin
        dbText "Name" ="UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EOLYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EOLCount"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =199
    Top =56
    Right =1626
    Bottom =1054
    Left =-1
    Top =-1
    Right =1409
    Bottom =527
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =493
        Top =0
        Name ="vw_sixkeys"
        Name =""
    End
End
