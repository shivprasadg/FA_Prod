Operation =1
Option =0
Begin InputTables
    Name ="vw_ATF_Avail_Signed_FAID"
End
Begin OutputColumns
    Expression ="vw_ATF_Avail_Signed_FAID.clientID"
    Expression ="vw_ATF_Avail_Signed_FAID.Clients_clientCompanyName"
    Expression ="vw_ATF_Avail_Signed_FAID.FAID"
    Expression ="vw_ATF_Avail_Signed_FAID.BDEX"
    Alias ="Schedules"
    Expression ="vw_ATF_Avail_Signed_FAID.Schedules"
    Alias ="Groups"
    Expression ="IIf([Schedules]>1,Max([GroupsIn])*[Schedules],Max([GroupsIn]))"
    Alias ="ShortGUID"
    Expression ="First(vw_ATF_Avail_Signed_FAID.ShortGUID)"
    Alias ="Schedule_SchShortDesc"
    Expression ="First(vw_ATF_Avail_Signed_FAID.Schedule_SchShortDesc)"
    Alias ="Spec_UnitQty"
    Expression ="Sum(vw_ATF_Avail_Signed_FAID.Spec_UnitQty)"
    Expression ="vw_ATF_Avail_Signed_FAID.SalesPersonID"
    Alias ="SPInitials"
    Expression ="vw_ATF_Avail_Signed_FAID.BDEX"
    Expression ="vw_ATF_Avail_Signed_FAID.ClientGroupId"
    Alias ="LPAFcomp"
    Expression ="\"2/5/2019\""
    Expression ="vw_ATF_Avail_Signed_FAID.GuarantorName"
    Expression ="vw_ATF_Avail_Signed_FAID.GuarantorID"
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_DealShortDesc"
    Expression ="vw_ATF_Avail_Signed_FAID.SchGrp_OrigEquipCost"
    Expression ="vw_ATF_Avail_Signed_FAID.UnitDesc"
    Alias ="ActualClose"
    Expression ="vw_ATF_Avail_Signed_FAID.[LPAFDate]"
    Expression ="vw_ATF_Avail_Signed_FAID.SF_ClientLegalName"
    Expression ="vw_ATF_Avail_Signed_FAID.Spec_Name"
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_dealSchedNum"
    Alias ="LPAFDate"
    Expression ="vw_ATF_Avail_Signed_FAID.LPAFDate"
    Expression ="vw_ATF_Avail_Signed_FAID.SFDC_Accounts_Id"
End
Begin OrderBy
    Expression ="vw_ATF_Avail_Signed_FAID.SF_ClientLegalName"
    Flag =0
End
Begin Groups
    Expression ="vw_ATF_Avail_Signed_FAID.clientID"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Clients_clientCompanyName"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.FAID"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Schedules"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.SalesPersonID"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.BDEX"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.ClientGroupId"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.GuarantorName"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.GuarantorID"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_DealShortDesc"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.SchGrp_OrigEquipCost"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.UnitDesc"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.[LPAFDate]"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.SF_ClientLegalName"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Spec_Name"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_dealSchedNum"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.LPAFDate"
    GroupLevel =0
    Expression ="vw_ATF_Avail_Signed_FAID.SFDC_Accounts_Id"
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
        dbText "Name" ="Groups"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ShortGUID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.Clients_clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule_SchShortDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.GuarantorName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Spec_UnitQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.BDEX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.GuarantorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.SalesPersonID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedules"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.Deals_DealShortDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4020"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SPInitials"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.UnitDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LPAFcomp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.SchGrp_OrigEquipCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.SF_ClientLegalName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.Spec_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.Deals_dealSchedNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ActualClose"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LPAFDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.SFDC_Accounts_Id"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-101
    Top =41
    Right =1025
    Bottom =904
    Left =-1
    Top =-1
    Right =1102
    Bottom =399
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =332
        Bottom =505
        Top =0
        Name ="vw_ATF_Avail_Signed_FAID"
        Name =""
    End
End
