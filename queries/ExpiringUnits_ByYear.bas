Operation =1
Option =0
Where ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Clients"
End
Begin OutputColumns
    Expression ="vw_SixKeys.ClientShortName"
    Alias ="BecameClient"
    Expression ="Clients.clientBecameDate"
    Expression ="vw_SixKeys.MLOrig"
    Alias ="UnitType"
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
    Expression ="vw_SixKeys.AssetID"
    Expression ="vw_SixKeys.Status"
    Alias ="OYear"
    Expression ="DatePart(\"yyyy\",CDate(Nz([OffLeaseDate],Nz([ProjEOL],\"12/31/2050\"))))"
    Alias ="O2013"
    Expression ="IIf([OYear]<=2013,1,0)"
    Alias ="O2014"
    Expression ="IIf([OYear]<=2014,1,0)"
    Alias ="O2015"
    Expression ="IIf([OYear]=2015,1,0)"
    Alias ="O2016"
    Expression ="IIf([OYear]=2016,1,0)"
    Alias ="O2017"
    Expression ="IIf([OYear]=2017,1,0)"
    Alias ="O2018"
    Expression ="IIf([OYear]=2018,1,0)"
    Alias ="O2019"
    Expression ="IIf([OYear]=2019,1,0)"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Clients"
    Expression ="vw_SixKeys.clientID = Clients.clientID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
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
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbInteger "ColumnWidth" ="2025"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BecameClient"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitType"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2019"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="O2013"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =74
    Top =2
    Right =1502
    Bottom =778
    Left =-1
    Top =-1
    Right =1410
    Bottom =365
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =217
        Bottom =637
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =286
        Top =116
        Right =430
        Bottom =571
        Top =0
        Name ="Clients"
        Name =""
    End
End
