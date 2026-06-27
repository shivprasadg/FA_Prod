Operation =1
Option =0
Having ="(((Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\""
    ",[Type],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler"
    "\" Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" A"
    "nd [SubType]=\"TBD\",\"Tractor TBD\"))<>\"\"))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="Awarded_Deals_Client"
End
Begin OutputColumns
    Alias ="Leased"
    Expression ="IIf([MLOrig]=\"FA\",\"FA\",\"Tracking\")"
    Alias ="FAYear"
    Expression ="Nz([DealYear],IIf([mlorig]=\"FA\",DatePart(\"yyyy\",Nz([BuildDate],Nz([ReportSta"
        "rtDate],[EsitmatedDeliveryDate]))),[MYear]-1))"
    Expression ="vw_SixKeys.Status"
    Alias ="UnitType"
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.UnitID)"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="Awarded_Deals_Client"
    Expression ="vw_SixKeys.FAID = Awarded_Deals_Client.FAID"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Awarded_Deals_Client"
    Expression ="vw_SixKeys.clientID = Awarded_Deals_Client.clientID"
    Flag =2
End
Begin OrderBy
    Expression ="IIf([MLOrig]=\"FA\",\"FA\",\"Tracking\")"
    Flag =0
    Expression ="Nz([DealYear],IIf([mlorig]=\"FA\",DatePart(\"yyyy\",Nz([BuildDate],Nz([ReportSta"
        "rtDate],[EsitmatedDeliveryDate]))),[MYear]-1))"
    Flag =1
    Expression ="vw_SixKeys.Status"
    Flag =0
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
    Flag =0
End
Begin Groups
    Expression ="IIf([MLOrig]=\"FA\",\"FA\",\"Tracking\")"
    GroupLevel =0
    Expression ="Nz([DealYear],IIf([mlorig]=\"FA\",DatePart(\"yyyy\",Nz([BuildDate],Nz([ReportSta"
        "rtDate],[EsitmatedDeliveryDate]))),[MYear]-1))"
    GroupLevel =0
    Expression ="vw_SixKeys.Status"
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
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="vw_SixKeys.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Leased"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAYear"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =41
    Top =117
    Right =1469
    Bottom =702
    Left =-1
    Top =-1
    Right =1410
    Bottom =332
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =361
        Bottom =647
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =464
        Top =86
        Right =608
        Bottom =315
        Top =0
        Name ="Awarded_Deals_Client"
        Name =""
    End
End
