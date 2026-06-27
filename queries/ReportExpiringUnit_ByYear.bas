Operation =1
Option =0
Where ="(((vw_SixKeys.UnitDescShortVer) Not Like '%WARRANTY%') AND ((vw_SixKeys.ProjEOL)"
    " Is Not Null))"
Having ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="UnitType"
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.AssetID)"
    Alias ="ExpireQtr"
    Expression ="DatePart(\"q\",[ProjEOL]) & \" Qtr\""
    Alias ="ExpireYear"
    Expression ="DatePart(\"yyyy\",[ProjEOL])"
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
    GroupLevel =0
    Expression ="DatePart(\"q\",[ProjEOL]) & \" Qtr\""
    GroupLevel =0
    Expression ="DatePart(\"yyyy\",[ProjEOL])"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
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
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="ExpireYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExpireQtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitType"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =146
    Right =1307
    Bottom =969
    Left =-1
    Top =-1
    Right =1201
    Bottom =529
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =356
        Bottom =493
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
End
