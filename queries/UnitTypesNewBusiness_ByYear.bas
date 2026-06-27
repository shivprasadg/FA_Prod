Operation =1
Option =0
Where ="(((vw_sixkeys.MLOrig)=\"FA\"))"
Having ="(((Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\""
    ",[Type],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler"
    "\" Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" A"
    "nd [SubType]=\"TBD\",\"Tractor TBD\")) Is Not Null))"
Begin InputTables
    Name ="vw_sixkeys"
    Name ="Clients"
End
Begin OutputColumns
    Alias ="Accepted"
    Expression ="DatePart(\"yyyy\",CDate(Nz(Nz([AcceptedDate],Nz([ActualDeliveryDate],[EsitmatedD"
        "eliveryDate])),Date())))"
    Alias ="NewClient"
    Expression ="DatePart(\"yyyy\",[clientBecameDate])"
    Alias ="UnitType"
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
    Alias ="NewBusiness"
    Expression ="Sum(IIf([NewClient]=[Accepted],1,0))"
End
Begin Joins
    LeftTable ="vw_sixkeys"
    RightTable ="Clients"
    Expression ="vw_sixkeys.clientID = Clients.clientID"
    Flag =1
End
Begin OrderBy
    Expression ="DatePart(\"yyyy\",CDate(Nz(Nz([AcceptedDate],Nz([ActualDeliveryDate],[EsitmatedD"
        "eliveryDate])),Date())))"
    Flag =0
    Expression ="DatePart(\"yyyy\",[clientBecameDate])"
    Flag =0
    Expression ="Switch([Type]=\"Yard Tractor\" Or [Type]=\"Straight Truck\" Or [Type]=\"Van\",[T"
        "ype],[SubType]=\"Daycab\" Or [SubType]=\"Sleeper\" Or [SubType]=\"Auto Hauler\" "
        "Or [subType]=\"Dolly\",[SubType],[Type]='Trailer',[Type],[Type]=\"Tractor\" And "
        "[SubType]=\"TBD\",\"Tractor TBD\")"
    Flag =0
End
Begin Groups
    Expression ="DatePart(\"yyyy\",CDate(Nz(Nz([AcceptedDate],Nz([ActualDeliveryDate],[EsitmatedD"
        "eliveryDate])),Date())))"
    GroupLevel =0
    Expression ="DatePart(\"yyyy\",[clientBecameDate])"
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
        dbText "Name" ="NewBusiness"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewClient"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Accepted"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =279
    Top =64
    Right =1706
    Bottom =1062
    Left =-1
    Top =-1
    Right =1409
    Bottom =561
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
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
End
