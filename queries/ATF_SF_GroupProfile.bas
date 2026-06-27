Operation =1
Option =2
Begin InputTables
    Name ="vw_ATF_Avail_Signed_FAID"
End
Begin OutputColumns
    Expression ="vw_ATF_Avail_Signed_FAID.ClientShNm"
    Expression ="vw_ATF_Avail_Signed_FAID.GroupsIn"
    Expression ="vw_ATF_Avail_Signed_FAID.SchName"
    Alias ="ATFGroupSpecDesc"
    Expression ="vw_ATF_Avail_Signed_FAID.GroupName"
    Alias ="UnitSpecsInGroup"
    Expression ="1"
    Alias ="UnitsInGroup"
    Expression ="vw_ATF_Avail_Signed_FAID.Spec_UnitQty"
    Alias ="GrpFAID"
    Expression ="vw_ATF_Avail_Signed_FAID.FAID"
    Alias ="ExchangeITGroup"
    Expression ="-1"
    Alias ="ExchangeITType"
    Expression ="'Specified Mo'"
    Alias ="ExchangeITMonth"
    Expression ="vw_ATF_Avail_Signed_FAID.SchGrp_ExchangeITMonth"
    Alias ="ExtendITOptionTerm"
    Expression ="12"
    Alias ="PrimaryLeaseTerm"
    Expression ="vw_ATF_Avail_Signed_FAID.SchGrp_LeaseTermPrimary"
    Alias ="PrimaryLeaseNotice"
    Expression ="3"
    Alias ="LeasePymtFreq"
    Expression ="'M'"
    Alias ="LeasePaidAdvArr"
    Expression ="'Adv'"
    Alias ="SwapIndex"
    Expression ="'Swap'"
    Alias ="SwapDate"
    Expression ="vw_ATF_Avail_Signed_FAID.SalesSwapDate"
    Alias ="SwapRate"
    Expression ="vw_ATF_Avail_Signed_FAID.SalesSwapRate"
    Alias ="SwapTermYears"
    Expression ="vw_ATF_Avail_Signed_FAID.SalesSwapTerm"
    Alias ="SwapYield"
    Expression ="vw_ATF_Avail_Signed_FAID.SchGrp_SgrpSalesYield"
    Alias ="SwapBonus"
    Expression ="vw_ATF_Avail_Signed_FAID.SchGrp_SGrpSalesBonus"
    Alias ="SwapAdjustment"
    Expression ="10*0.01"
    Alias ="ResidualAvg"
    Expression ="CCur([Units_UnitSalesResid])"
    Alias ="UATPerUnit"
    Expression ="CCur([UTA_Amount_per_Unit__c])"
    Alias ="PLB"
    Expression ="0"
    Alias ="LeaseType"
    Expression ="'CAPITAL'"
    Alias ="BLCDBasisID"
    Expression ="1"
    Alias ="PerDiemBasisID"
    Expression ="2"
    Alias ="EstDelDate"
    Expression ="vw_ATF_Avail_Signed_FAID.Deals_dealEstDelDate"
    Alias ="NumUTAs"
    Expression ="vw_ATF_Avail_Signed_FAID.Total_UTAs_per_ERA__c"
    Expression ="vw_ATF_Avail_Signed_FAID.MilesPerYear"
    Expression ="vw_ATF_Avail_Signed_FAID.HoursPerYear"
    Alias ="Fee_Projected"
    Expression ="CCur([Fee_Dollar__c])*[Spec_UnitQty]"
    Alias ="Bonus"
    Expression ="vw_ATF_Avail_Signed_FAID.SchGrp_SGrpSalesBonus"
End
Begin OrderBy
    Expression ="vw_ATF_Avail_Signed_FAID.ClientShNm"
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
        dbText "Name" ="NumUTAs"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="EstDelDate"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.ClientShNm"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.SchName"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.GroupsIn"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitsInGroup"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="0"
    End
    Begin
        dbText "Name" ="SwapIndex"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="PrimaryLeaseTerm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="ATFGroupSpecDesc"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitSpecsInGroup"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="GrpFAID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="ExchangeITType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="ExchangeITMonth"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="LeasePymtFreq"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="SwapDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="SwapRate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="SwapYield"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="SwapBonus"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="PerDiemBasisID"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Leasetype"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLCDBasisID"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SwapAdjustment"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ResidualAvg"
        dbInteger "ColumnWidth" ="2385"
        dbInteger "ColumnOrder" ="25"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PLB"
        dbInteger "ColumnOrder" ="29"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExtendITOptionTerm"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PrimaryLeaseNotice"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SwapTermYears"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UATPerUnit"
        dbInteger "ColumnOrder" ="26"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchangeITGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.HoursPerYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_ATF_Avail_Signed_FAID.MilesPerYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fee_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bonus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =154
    Top =98
    Right =1015
    Bottom =943
    Left =-1
    Top =-1
    Right =837
    Bottom =242
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =34
        Top =0
        Right =266
        Bottom =372
        Top =0
        Name ="vw_ATF_Avail_Signed_FAID"
        Name =""
    End
End
