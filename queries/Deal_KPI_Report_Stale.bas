Operation =1
Option =2
Where ="(((Schedule_First_BLCDate.ClientID) Is Null))"
Having ="(((Deals.dealOutcome)=\"Awarded\") AND ((Deals.dealEnteredDate)>#12/31/2012#))"
Begin InputTables
    Name ="Deals"
    Name ="Clients"
    Name ="Schedule_First_BLCDate"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="Clients.ClientShNm"
    Alias ="DealDesc"
    Expression ="Deals.dealShortDesc"
    Alias ="DealResult"
    Expression ="Deals.dealOutcome"
    Alias ="DealEntered"
    Expression ="Deals.dealEnteredDate"
    Alias ="ResultDate"
    Expression ="Deals.dealOutcomeDate"
    Alias ="DaysToAward"
    Expression ="DateDiff('d',[dealEnteredDate],[dealOutcomeDate])"
    Alias ="Schedule"
    Expression ="First(Schedule_First_BLCDate.Schedule)"
    Alias ="ScheduleAdded"
    Expression ="Schedule_First_BLCDate.DateSchAdded"
    Alias ="StaleDays"
    Expression ="DateDiff('d',[dealOutcomeDate],Date())"
    Alias ="Assignee"
    Expression ="First(Schedule_First_BLCDate.Assignee)"
    Alias ="BLCD"
    Expression ="First(Schedule_First_BLCDate.FirstBLCD)"
    Alias ="OPSDays"
    Expression ="DateDiff('d',[DateSchAdded],First([FirstBLCD]))"
End
Begin Joins
    LeftTable ="Deals"
    RightTable ="Clients"
    Expression ="Deals.clientRef = Clients.clientID"
    Flag =2
    LeftTable ="Clients"
    RightTable ="Schedule_First_BLCDate"
    Expression ="Clients.clientID = Schedule_First_BLCDate.ClientID"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="Deals.dealEnteredDate"
    Flag =1
End
Begin Groups
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="Deals.dealShortDesc"
    GroupLevel =0
    Expression ="Deals.dealOutcome"
    GroupLevel =0
    Expression ="Deals.dealEnteredDate"
    GroupLevel =0
    Expression ="Deals.dealOutcomeDate"
    GroupLevel =0
    Expression ="Schedule_First_BLCDate.DateSchAdded"
    GroupLevel =0
    Expression ="Deals.dealEnteredDate"
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
        dbText "Name" ="OPSDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StaleDays"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =43
    Top =-6
    Right =1753
    Bottom =967
    Left =-1
    Top =-1
    Right =1678
    Bottom =491
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
        Left =485
        Top =154
        Right =629
        Bottom =379
        Top =0
        Name ="Schedule_First_BLCDate"
        Name =""
    End
End
