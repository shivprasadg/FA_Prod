Operation =1
Option =2
Having ="(((Deals.dealOutcome)=\"Awarded\") AND ((Min(Deals.dealEnteredDate))>#12/31/2012"
    "#))"
Begin InputTables
    Name ="Deals"
    Name ="Clients"
    Name ="Schedule_First_BLCDate"
    Name ="ATFMaxRevNo"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="Clients.ClientShNm"
    Alias ="DealDesc"
    Expression ="Deals.dealShortDesc"
    Alias ="DealResult"
    Expression ="Deals.dealOutcome"
    Alias ="DealEntered"
    Expression ="Min(Deals.dealEnteredDate)"
    Alias ="ResultDate"
    Expression ="Deals.dealOutcomeDate"
    Alias ="DaysToAward"
    Expression ="DateDiff('d',Min([dealEnteredDate]),[dealOutcomeDate])"
    Alias ="Schedule"
    Expression ="Min(Schedule_First_BLCDate.Schedule)"
    Alias ="ScheduleAdded"
    Expression ="IIf(IsNull(Min([DateSchAdded])),\"PENDING\",Min([DateSchAdded]))"
    Alias ="LPAFDays"
    Expression ="IIf(IsNull(Min([DateSchAdded])),DateDiff('d',[dealOutcomeDate],Date()),DateDiff("
        "'d',[dealOutcomeDate],Min([DateSchAdded])))"
    Alias ="Assignee"
    Expression ="Min(Schedule_First_BLCDate.Assignee)"
    Alias ="BLCD"
    Expression ="Nz(Min([FirstBLCD]),\"PENDING\")"
    Alias ="OPSDays"
    Expression ="IIf(IsNull(Min([FirstBLCD])),DateDiff('d',Min([DateSchAdded]),Date()),DateDiff('"
        "d',Min([DateSchAdded]),Min([FirstBLCD])))"
    Alias ="ATFRev"
    Expression ="Nz([ATFMaxRevNo].[RevDt],\"n/a\")"
End
Begin Joins
    LeftTable ="Deals"
    RightTable ="Clients"
    Expression ="Deals.clientRef = Clients.clientID"
    Flag =1
    LeftTable ="Deals"
    RightTable ="Schedule_First_BLCDate"
    Expression ="Deals.FAID = Schedule_First_BLCDate.FAID"
    Flag =2
    LeftTable ="Deals"
    RightTable ="Schedule_First_BLCDate"
    Expression ="Deals.clientRef = Schedule_First_BLCDate.ClientID"
    Flag =2
    LeftTable ="Schedule_First_BLCDate"
    RightTable ="ATFMaxRevNo"
    Expression ="Schedule_First_BLCDate.ScheduleID = ATFMaxRevNo.SchID"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="Min(Deals.dealEnteredDate)"
    Flag =1
End
Begin Groups
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="Deals.dealShortDesc"
    GroupLevel =0
    Expression ="Deals.dealOutcome"
    GroupLevel =0
    Expression ="Deals.dealOutcomeDate"
    GroupLevel =0
    Expression ="Nz([ATFMaxRevNo].[RevDt],\"n/a\")"
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
        dbText "Name" ="DealEntered"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="DealResult"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="DealDesc"
        dbInteger "ColumnWidth" ="6960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="ResultDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="ScheduleAdded"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysToAward"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LPAFDays"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="OPSDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFRev"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =95
    Top =81
    Right =1899
    Bottom =874
    Left =-1
    Top =-1
    Right =1772
    Bottom =338
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =233
        Bottom =526
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =270
        Top =32
        Right =414
        Bottom =377
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =515
        Top =195
        Right =659
        Bottom =420
        Top =0
        Name ="Schedule_First_BLCDate"
        Name =""
    End
    Begin
        Left =691
        Top =240
        Right =835
        Bottom =384
        Top =0
        Name ="ATFMaxRevNo"
        Name =""
    End
End
