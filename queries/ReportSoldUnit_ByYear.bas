Operation =1
Option =0
Where ="(((vw_SixKeys.UnitDescShortVer) Not Like '%WARRANTY%'))"
Having ="(((DatePart(\"yyyy\",[InvoicePaidDate])) Is Not Null))"
Begin InputTables
    Name ="vw_SixKeys"
    Name ="vw_Units_AutoHaulerPaired"
    Name ="Rmkt"
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
    Alias ="SoldQtr"
    Expression ="DatePart(\"q\",[InvoicePaidDate]) & \" Qtr\""
    Alias ="SoldYear"
    Expression ="DatePart(\"yyyy\",[InvoicePaidDate])"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="vw_Units_AutoHaulerPaired"
    Expression ="vw_SixKeys.UnitID = vw_Units_AutoHaulerPaired.TractorUnitID"
    Flag =2
    LeftTable ="vw_SixKeys"
    RightTable ="Rmkt"
    Expression ="vw_SixKeys.UnitID = Rmkt.UnitRef"
    Flag =1
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
    Expression ="DatePart(\"q\",[InvoicePaidDate]) & \" Qtr\""
    GroupLevel =0
    Expression ="DatePart(\"yyyy\",[InvoicePaidDate])"
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
dbMemo "Filter" ="([ReportSoldUnit_ByYear].[UnitType]=\"Daycab\")"
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
        dbText "Name" ="UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SoldYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SoldQtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitDescShortVer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExpireQtr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.offleasedate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_Units_AutoHaulerPaired.TractorUnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExpireYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rmkt.InvoicePaidDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.MLOrig"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =126
    Top =26
    Right =1345
    Bottom =849
    Left =-1
    Top =-1
    Right =1201
    Bottom =512
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
    Begin
        Left =453
        Top =107
        Right =597
        Bottom =251
        Top =0
        Name ="vw_Units_AutoHaulerPaired"
        Name =""
    End
    Begin
        Left =449
        Top =266
        Right =593
        Bottom =507
        Top =0
        Name ="Rmkt"
        Name =""
    End
End
