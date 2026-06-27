Operation =1
Option =0
Begin InputTables
    Name ="vw_ATF_Avail_Signed_FAID"
    Name ="Salesperson"
End
Begin OutputColumns
    Alias ="FAID"
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_FAID"
    Alias ="SumOfDeals_dealGrandTotal"
    Expression ="Sum(vw_ATF_Avail_Signed_FAID.Deals_dealGrandTotal)"
    Alias ="dealShortDesc"
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_DealShortDesc"
    Alias ="dealRFPDate"
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_RFPDate"
    Alias ="dealLPAFDue"
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_dealLPAFDue"
    Alias ="dealLPAFcomp"
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_dealLPAFcomp"
    Alias ="dealUnitQty"
    Expression ="Sum(vw_ATF_Avail_Signed_FAID.Deals_dealUnitQty)"
    Alias ="clientRef"
    Expression ="[Forms]![ATF_SFImporter]![LesseeID]"
    Alias ="dealOutcome"
    Expression ="'Awarded'"
    Alias ="dealOutcomeDate"
    Expression ="CDate([LPAFDate])"
    Alias ="dealEnteredDate"
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_dealDateEntered"
    Alias ="dealEstDelDate"
    Expression ="Max(vw_ATF_Avail_Signed_FAID.Deals_dealEstDelDate)"
    Alias ="dealMPY"
    Expression ="Min(vw_ATF_Avail_Signed_FAID.Deals_dealMPY)"
    Alias ="dealTerm"
    Expression ="Max(vw_ATF_Avail_Signed_FAID.Deals_dealTerm)"
    Alias ="dealUnitPrice"
    Expression ="Avg(vw_ATF_Avail_Signed_FAID.Deals_dealUnitPrice)"
    Alias ="dealGrandTotal"
    Expression ="Sum(vw_ATF_Avail_Signed_FAID.Deals_dealGrandTotal)"
    Alias ="DealLOEx"
    Expression ="Salesperson.SPInitials"
    Expression ="vw_ATF_Avail_Signed_FAID.Type"
End
Begin Joins
    LeftTable ="vw_ATF_Avail_Signed_FAID"
    RightTable ="Salesperson"
    Expression ="vw_ATF_Avail_Signed_FAID.SalespersonID = Salesperson.SalespersonID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_FAID"
    Flag =0
End
Begin Groups
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_FAID"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_DealShortDesc"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_RFPDate"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_dealLPAFDue"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_dealLPAFcomp"
    GroupLevel =0
    Expression ="CDate([LPAFDate])"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_dealDateEntered"
    GroupLevel =0
    Expression ="Salesperson.SPInitials"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Type"
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
        dbText "Name" ="dealUnitPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealEnteredDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealGrandTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealOutcomeDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="dealShortDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealRFPDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealLPAFDue"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="dealLPAFcomp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealUnitQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dealOutcome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DealLOEx"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfDeals_dealGrandTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.Type"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =27
    Top =0
    Right =1726
    Bottom =732
    Left =-1
    Top =-1
    Right =1681
    Bottom =324
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =268
        Top =35
        Right =737
        Bottom =310
        Top =0
        Name ="vw_ATF_Avail_Signed_FAID"
        Name =""
    End
    Begin
        Left =850
        Top =163
        Right =1093
        Bottom =307
        Top =0
        Name ="Salesperson"
        Name =""
    End
End
