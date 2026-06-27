Operation =1
Option =0
Where ="(((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Clients"
    Name ="UnitExitOption"
    Name ="UnitExitOption"
    Alias ="UnitExitOption_1"
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
    Alias ="Accepted"
    Expression ="CDate(Nz(Nz([AcceptedDate],Nz([ActualDeliveryDate],[EsitmatedDeliveryDate])),Dat"
        "e()))"
    Expression ="vw_SixKeys.BLCD"
    Expression ="UnitExitOption.EndOrActualDate"
    Alias ="OffLease"
    Expression ="CDate(Nz([UnitExitOption].[EndOrActualDate],Nz([OffLeaseDate],Nz([ProjEOL],\"12/"
        "31/2050\"))))"
    Expression ="vw_SixKeys.ProjEOL"
    Alias ="ExchangeIT"
    Expression ="IIf([UnitExitOption].[OptionTypeID]=1,CDate(Format([UnitExitOption].[DatePosted]"
        ",\"mm/dd/yyyy\")))"
    Expression ="UnitExitOption.ReplacedByGroupID"
    Alias ="Extended"
    Expression ="IIf([UnitExitOption].[OptionTypeID]=3,CDate(Format([UnitExitOption].[DatePosted]"
        ",\"mm/dd/yyyy\")))"
    Alias ="DYear"
    Expression ="DatePart(\"yyyy\",CDate(Nz(Nz([AcceptedDate],Nz([ActualDeliveryDate],[EsitmatedD"
        "eliveryDate])),Date())))"
    Alias ="BCYear"
    Expression ="DatePart(\"yyyy\",[clientBecameDate])"
    Alias ="OYear"
    Expression ="CInt(DatePart(\"yyyy\",Nz([OffLease],CDate(\"1/1/2050\"))))"
    Alias ="New2014"
    Expression ="IIf([BCYear]=2014 And [DYear]=[BCYear],1,0)"
    Alias ="New2015"
    Expression ="IIf([BCYear]=2015 And [DYear]=[BCYear],1,0)"
    Alias ="New2016"
    Expression ="IIf([BCYear]=2016 And [DYear]=[BCYear],1,0)"
    Alias ="New2017"
    Expression ="IIf([BCYear]=2017 And [DYear]=[BCYear],1,0)"
    Alias ="New2018"
    Expression ="IIf([BCYear]=2018 And [DYear]=[BCYear],1,0)"
    Alias ="New2019"
    Expression ="IIf([BCYear]=2019 And [DYear]=[BCYear],1,0)"
    Alias ="A2014"
    Expression ="IIf([DYear]<=2014 And [OYear]>=2014,1,0)"
    Alias ="A2015"
    Expression ="IIf([DYear]<=2015 And [OYear]>=2015,1,0)"
    Alias ="A2016"
    Expression ="IIf([DYear]<=2016 And [OYear]>=2016,1,0)"
    Alias ="A2017"
    Expression ="IIf([DYear]<=2017 And [OYear]>=2017,1,0)"
    Alias ="A2018"
    Expression ="IIf([DYear]<=2018 And [OYear]>=2018,1,0)"
    Alias ="PartOfEXchangeIT"
    Expression ="IIf(IsNull([UnitExitOption_1].[ReplacedByGroupID]),0,1)"
    Alias ="Added2014"
    Expression ="IIf([DYear]=2014 And [DYear]<>[BCYear],1,0)"
    Alias ="Added2015"
    Expression ="IIf([DYear]=2015 And [DYear]<>[BCYear],1,0)"
    Alias ="Added2016"
    Expression ="IIf([DYear]=2016 And [DYear]<>[BCYear],1,0)"
    Alias ="Added2017"
    Expression ="IIf([DYear]=2017 And [DYear]<>[BCYear],1,0)"
    Alias ="Added2018"
    Expression ="IIf([DYear]=2018 And [DYear]<>[BCYear],1,0)"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Clients"
    Expression ="vw_SixKeys.clientID = Clients.clientID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="UnitExitOption"
    Expression ="vw_SixKeys.UnitID = UnitExitOption.UnitID"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="UnitExitOption_1"
    Expression ="vw_SixKeys.SchID = UnitExitOption_1.ReplacedByGroupID"
    Flag =2
End
Begin OrderBy
    Expression ="CDate(Nz(Nz([AcceptedDate],Nz([ActualDeliveryDate],[EsitmatedDeliveryDate])),Dat"
        "e()))"
    Flag =0
    Expression ="DatePart(\"yyyy\",CDate(Nz(Nz([AcceptedDate],Nz([ActualDeliveryDate],[EsitmatedD"
        "eliveryDate])),Date())))"
    Flag =0
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
dbMemo "Filter" ="([UnitsByYear].[Added2014]<>0 Or [UnitsByYear].[Added2014] IS Null)"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.ClientShortName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="OYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="New2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BCYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="New2015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="New2016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="New2017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="New2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="New2019"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Accepted"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="BecameClient"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.BLCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="ExchangeIT"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ProjEOL"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByGroupID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Extended"
        dbInteger "ColumnWidth" ="1395"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OffLease"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.EndOrActualDate"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PartOfEXchangeIT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitType"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Added2015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Added2016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Added2017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Added2018"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Added2014"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =160
    Top =143
    Right =1588
    Bottom =1136
    Left =-1
    Top =-1
    Right =1410
    Bottom =586
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
    Begin
        Left =537
        Top =141
        Right =753
        Bottom =656
        Top =0
        Name ="UnitExitOption"
        Name =""
    End
    Begin
        Left =809
        Top =67
        Right =1082
        Bottom =441
        Top =0
        Name ="UnitExitOption_1"
        Name =""
    End
End
