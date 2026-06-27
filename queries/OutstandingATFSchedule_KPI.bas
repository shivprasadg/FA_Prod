Operation =1
Option =2
Having ="(((Deals.dealOutcome)=\"Awarded\") AND ((Max(DateValue([dealEnteredDate])))>#12/"
    "31/2012#) AND ((Deals.ExcludeFromKPI) Is Null Or (Deals.ExcludeFromKPI)=0))"
Begin InputTables
    Name ="Deals"
    Name ="Clients"
    Name ="Schedule_Group_Entered_KPI"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="Clients.ClientShNm"
    Alias ="DealDesc"
    Expression ="Deals.dealShortDesc"
    Alias ="DealResult"
    Expression ="Deals.dealOutcome"
    Alias ="DealEntered"
    Expression ="Max(DateValue([dealEnteredDate]))"
    Alias ="ResultDate"
    Expression ="Deals.dealOutcomeDate"
    Alias ="DaysToAward"
    Expression ="DateDiff('d',Min([dealEnteredDate]),[dealOutcomeDate])"
    Alias ="Schedule"
    Expression ="Min(Schedule_Group_Entered_KPI.Schedule)"
    Alias ="ScheduleAdded"
    Expression ="IIf(IsNull(Min([DateSchAdded])),\"PENDING\",Min([DateSchAdded]))"
    Alias ="Group"
    Expression ="Min(Schedule_Group_Entered_KPI.Group)"
    Alias ="GroupAdded"
    Expression ="IIf(IsNull(Min([DateGroupAdded])),\"PENDING\",Min([DateGroupAdded]))"
    Alias ="LPAFDays"
    Expression ="IIf(IsNull(Min([DateGroupAdded])),DateDiff('d',[dealOutcomeDate],Date()),DateDif"
        "f('d',[dealOutcomeDate],Min([DateGroupAdded])))"
    Alias ="BLCD"
    Expression ="Nz(Min([Schedule_Group_Entered_KPI].[BLCD]),\"PENDING\")"
    Alias ="OPSDays"
    Expression ="IIf(IsNull(Min([Schedule_Group_Entered_KPI].[ApprovalRequest])),DateDiff('d',Min"
        "([DateSchAdded]),Date()),DateDiff('d',Min([DateSchAdded]),Min([Schedule_Group_En"
        "tered_KPI].[ApprovalRequest])))"
    Expression ="Deals.FAID"
    Expression ="Deals.dealUnitQty"
    Expression ="Deals.dealGrandTotal"
    Expression ="Schedule_Group_Entered_KPI.ApprovalRequest"
    Expression ="Schedule_Group_Entered_KPI.ApprovalReply"
    Expression ="Deals.ExcludeFromKPI"
End
Begin Joins
    LeftTable ="Deals"
    RightTable ="Clients"
    Expression ="Deals.clientRef = Clients.clientID"
    Flag =1
    LeftTable ="Deals"
    RightTable ="Schedule_Group_Entered_KPI"
    Expression ="Deals.FAID = Schedule_Group_Entered_KPI.FAID"
    Flag =2
    LeftTable ="Deals"
    RightTable ="Schedule_Group_Entered_KPI"
    Expression ="Deals.clientRef = Schedule_Group_Entered_KPI.ClientID"
    Flag =2
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="Max(DateValue([dealEnteredDate]))"
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
    Expression ="Deals.FAID"
    GroupLevel =0
    Expression ="Deals.dealUnitQty"
    GroupLevel =0
    Expression ="Deals.dealGrandTotal"
    GroupLevel =0
    Expression ="Schedule_Group_Entered_KPI.ApprovalRequest"
    GroupLevel =0
    Expression ="Schedule_Group_Entered_KPI.ApprovalReply"
    GroupLevel =0
    Expression ="Deals.ExcludeFromKPI"
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
        dbInteger "ColumnOrder" ="5"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="DealResult"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="DealDesc"
        dbInteger "ColumnWidth" ="6960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="ResultDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="ScheduleAdded"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DaysToAward"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LPAFDays"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="OPSDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.FAID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealGrandTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealUnitQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupAdded"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule_Group_Entered_KPI.ApprovalReply"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule_Group_Entered_KPI.ApprovalRequest"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.ExcludeFromKPI"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =112
    Top =26
    Right =1871
    Bottom =937
    Left =-1
    Top =-1
    Right =1727
    Bottom =390
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =33
        Top =-15
        Right =218
        Bottom =447
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
        Left =488
        Top =78
        Right =647
        Bottom =332
        Top =0
        Name ="Schedule_Group_Entered_KPI"
        Name =""
    End
End
