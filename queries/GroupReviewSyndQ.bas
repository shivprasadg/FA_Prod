Operation =1
Option =2
Having ="(((SchGrp.SGrpID)=[forms]![groupform]![sgrpid]))"
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="FSL"
    Name ="SchGrp"
    Name ="Units"
    Name ="UnitGroupPerdiemsSyndTotals"
    Name ="Clients"
    Alias ="Clients_1"
    Name ="Banks"
    Name ="UnitTotalOEC"
    Name ="SyndGroupStripsUTA_Summed"
    Name ="InterimFunding_Days_Rate_Group"
    Name ="PerDiem_Days_Rate_Group"
    Name ="UnitLevelTaxes"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="Clients.clientCompanyName"
    Alias ="ClientShortName"
    Expression ="Clients.ClientShNm"
    Alias ="MasterLease"
    Expression ="MstrLease.MLNo"
    Alias ="Guarantor"
    Expression ="Clients_1.clientCompanyName"
    Expression ="Clients.clientCompanyName"
    Alias ="Schedule"
    Expression ="Schedule.SchNo"
    Expression ="Schedule.SchID"
    Alias ="GroupID"
    Expression ="SchGrp.SGrpID"
    Expression ="SchGrp.SGrpNO"
    Alias ="GroupNum"
    Expression ="SchGrp.SGrpNO"
    Expression ="SchGrp.UnitsInGroup"
    Alias ="UnitCount"
    Expression ="Count([Units].[UnitId])*1"
    Expression ="SchGrp.OrigEquipCost"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.LeaseTermPrimary"
    Expression ="SchGrp.LeasePymtForGroup"
    Expression ="SchGrp.Assignee_Bank"
    Alias ="Assignee"
    Expression ="Banks.BankName"
    Expression ="SchGrp.Assigned_Indicator"
    Expression ="SchGrp.AssignmentDate"
    Expression ="SchGrp.FundDate"
    Expression ="SchGrp.SalePrice"
    Expression ="SchGrp.SGrpFSLPmt"
    Expression ="SchGrp.SGrpFSLSch"
    Expression ="SchGrp.LeasePaidAdvArr"
    Expression ="SchGrp.LeasePymtFreq"
    Expression ="SchGrp.ExchangeITGroup"
    Expression ="SchGrp.ExchangeITMonth"
    Expression ="SchGrp.ExchangeITType"
    Alias ="Bonus"
    Expression ="SchGrp.SGrpBonus"
    Alias ="Index"
    Expression ="SchGrp.SGrpIndex"
    Alias ="IndexDate"
    Expression ="SchGrp.SGrpIndDt"
    Alias ="IndexYear"
    Expression ="SchGrp.SGprIndYr"
    Expression ="SchGrp.IndexRate"
    Alias ="Yield"
    Expression ="SchGrp.SgrpYield"
    Expression ="SchGrp.EarlyTermNote"
    Expression ="SchGrp.SalesFSL"
    Alias ="ScheduleDesc"
    Expression ="Schedule.SchShortDesc"
    Alias ="UnitResidual"
    Expression ="Units.UnitSynResid"
    Alias ="SalesUnitResidual"
    Expression ="Avg(Units.unitsalesresid)"
    Alias ="UnitRVPct"
    Expression ="IIf([UnitSynResid]>0,Round([UnitSynResid]/[OrigEquipCost],3),0)"
    Expression ="SchGrp.Fee_Projected"
    Alias ="SalesBonus"
    Expression ="SchGrp.SGrpSalesBonus"
    Alias ="SalesIndex"
    Expression ="SchGrp.SGrpSalesIndex"
    Alias ="SalesIndexDate"
    Expression ="SchGrp.SGrpSalesIndDt"
    Alias ="SalesIndexYear"
    Expression ="SchGrp.SGprSalesIndYr"
    Alias ="SalesIndexRate"
    Expression ="SchGrp.SgrpSalesIndRt"
    Alias ="SalesYield"
    Expression ="SchGrp.SgrpSalesYield"
    Alias ="Spread"
    Expression ="[sgrpyield]-[IndexRate]"
    Alias ="UnitRentPayment"
    Expression ="CCur([unitrent])"
    Expression ="UnitTotalOEC.UnitLRF"
    Expression ="Schedule.SchFSL"
    Expression ="FSL.FSLYR1"
    Expression ="SchGrp.SaleDate"
    Expression ="SchGrp.SgrpTermPenalty"
    Expression ="SchGrp.SgrpSynPenalty"
    Expression ="SchGrp.WithHeldPymts"
    Expression ="SyndGroupStripsUTA_Summed.UTAPer"
    Alias ="RebateUnitActual"
    Expression ="SchGrp.SgrpSynRebate"
    Alias ="SyndPenalty"
    Expression ="SchGrp.SgrpSynPenalty"
    Alias ="StripsNum"
    Expression ="SyndGroupStripsUTA_Summed.GroupStrips"
    Alias ="StripsAmount"
    Expression ="SyndGroupStripsUTA_Summed.GroupStripAmount"
    Alias ="ServiceFee"
    Expression ="SchGrp.SGrpFSLPmt"
    Alias ="PayFreq"
    Expression ="Switch([LeasePymtFreq]=\"M\",\"Monthly\",[LeasePymtFreq]=\"Q\",\"Quarterly\",[Le"
        "asePymtFreq]=\"A\",\"Annualy\",[LeasePymtFreq]=\"S\",\"Semi-Annuly\")"
    Alias ="AdvanceArrears"
    Expression ="Switch([LeasePaidAdvArr]=\"Adv\",\"Advance\",[LeasePaidAdvArr]=\"Arr\",\"Arrears"
        "\")"
    Alias ="FeeTerm"
    Expression ="IIf(Nz([ScheduleGroupPaymentTerm],0)=0,[LeaseTermPrimary],[ScheduleGroupPaymentT"
        "erm])"
    Alias ="FullServiceFee"
    Expression ="((IIf(Nz([ScheduleGroupPaymentTerm],0)=0,[LeaseTermPrimary],[ScheduleGroupPaymen"
        "tTerm]))*[SGrpFSLPmt])*[Unitsingroup]"
    Alias ="PVFeeRate"
    Expression ="CDbl(Nz(forms!groupform![PVRate],0.05))"
    Alias ="ServFeePV"
    Expression ="IIf([forms]![groupform]![PVFeesYN]=True And [forms]![groupform]![ServFeePassTrou"
        "ghYN]=False,Round(PV(CDbl((Nz([PVFeeRate],0.05)/12)),[FeeTerm],([SGrpFSLPmt]*[Un"
        "itsInGroup])*-1,0,1),2),0)"
    Expression ="InterimFunding_Days_Rate_Group.InterimIntRate"
    Alias ="InterimLoanDays"
    Expression ="InterimFunding_Days_Rate_Group.AvgInterimLoanDays"
    Alias ="InterimFundingExpense"
    Expression ="InterimFunding_Days_Rate_Group.GroupInterimExp"
    Alias ="FullUserName"
    Expression ="FAUserFullName()"
    Expression ="PerDiem_Days_Rate_Group.AvgPerDiem"
    Expression ="PerDiem_Days_Rate_Group.PDDays"
    Alias ="PerDiemToDisplay"
    Expression ="IIf([SchGrp].[PerDiemActual]>0,[SchGrp].[PerDiemActual]*[UnitsInGroup],IIf([PerD"
        "iemGroup]>0,[PerDiemGroup],[PerDiemProj]))"
    Alias ="LRF"
    Expression ="CDbl(Round([LeasePymtForGroup]/[OrigEquipCost],7))"
    Alias ="MPY"
    Expression ="IIf([AllowedMilesPeriod]='P/Year',[SGprAllow1],([SGprAllow1]/([LeaseTermPrimary]"
        "-CInt(Nz([LeaseTermNotice],0)))*12))"
    Alias ="CostPerUnit"
    Expression ="CCur([OrigEquipCost]/[unitsingroup])"
    Alias ="FeeUnit"
    Expression ="IIf([SalePrice]>0,(([SalePrice]/[UnitsInGroup])-[UnitTotalCost]),0)"
    Alias ="FeeUnitCount"
    Expression ="[FeeUnit]*[UnitCount]"
    Alias ="FeeGroup"
    Expression ="CCur([SalePrice])-([Forms]![GroupForm]![CalcOEC])"
    Expression ="SchGrp.LoanDate"
    Alias ="GroupResidual"
    Expression ="Sum(Units.UnitSynResid)"
    Alias ="TotalCost"
    Expression ="SchGrp.origEquipCost"
    Expression ="UnitTotalOEC.UnitTotalCost"
    Expression ="SchGrp.LeasePymtForGroup"
    Alias ="LeasePymtUnit"
    Expression ="CCur([UnitRent])"
    Alias ="Group"
    Expression ="SchGrp.SGrpNO"
    Alias ="HBConversion"
    Expression ="CCur(Nz([HoldbackConversion],0))"
    Alias ="HBResidual"
    Expression ="CCur(Nz([HoldbackResidual],0))"
    Alias ="WithHeldNum"
    Expression ="CDbl(Nz([WithHeldPymts],0))"
    Alias ="WithHeldAmount"
    Expression ="[LeasePymtForGroup]*[WithHeldPymts]"
    Expression ="SchGrp.LastUpdatedDate"
    Expression ="SchGrp.LastUpdatedBy"
    Expression ="SchGrp.ScheduleGroupPaymentTerm"
    Expression ="SchGrp.SGrpFSLPmt"
    Expression ="SchGrp.SGrpRSLTPmt"
    Expression ="Schedule.SchShortDesc"
    Expression ="SchGrp.SyndMiscLabel1"
    Expression ="SchGrp.SyndMiscValue1"
    Expression ="SchGrp.SyndMiscLabel2"
    Expression ="SchGrp.SyndMiscValue2"
    Alias ="SynMemo"
    Expression ="First(SchGrp.SynMemo)"
    Alias ="SalesRent"
    Expression ="Units.UnitSRent"
    Expression ="SchGrp.RebateUnitProjected"
    Alias ="Rent"
    Expression ="Units.UnitRent"
    Alias ="AncillaryFee"
    Expression ="UnitLevelTaxes.FeeTaxable"
    Alias ="UnitTax"
    Expression ="CCur(Nz([UnitTaxes],0))"
    Alias ="LocState"
    Expression ="IIf([UnitTaxes]>0,[UnitlevelTaxes].[LState],\"-\")"
    Alias ="LTaxRate"
    Expression ="IIf([UnitTaxes]>0,[UnitlevelTaxes].[LTaxRate],0)"
    Alias ="GroupDesc"
    Expression ="SchGrp.SgrpDesc"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="FSL"
    Expression ="Schedule.FSLSCH = FSL.ID"
    Flag =2
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="Units"
    Expression ="SchGrp.SGrpID = Units.GroupREF"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="UnitGroupPerdiemsSyndTotals"
    Expression ="SchGrp.SGrpID = UnitGroupPerdiemsSyndTotals.GroupID"
    Flag =2
    LeftTable ="MstrLease"
    RightTable ="Clients_1"
    Expression ="MstrLease.MLGuarantor = Clients_1.clientID"
    Flag =2
    LeftTable ="SchGrp"
    RightTable ="Banks"
    Expression ="SchGrp.Assignee_Bank = Banks.BankShortName"
    Flag =2
    LeftTable ="Units"
    RightTable ="UnitTotalOEC"
    Expression ="Units.UnitID = UnitTotalOEC.AssetID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="SyndGroupStripsUTA_Summed"
    Expression ="SchGrp.SGrpID = SyndGroupStripsUTA_Summed.GroupID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="InterimFunding_Days_Rate_Group"
    Expression ="SchGrp.SGrpID = InterimFunding_Days_Rate_Group.GroupREF"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="PerDiem_Days_Rate_Group"
    Expression ="SchGrp.SGrpID = PerDiem_Days_Rate_Group.GroupREF"
    Flag =1
    LeftTable ="Units"
    RightTable ="UnitLevelTaxes"
    Expression ="Units.UnitID = UnitLevelTaxes.UnitID"
    Flag =1
End
Begin Groups
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Clients_1.clientCompanyName"
    GroupLevel =0
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="SchGrp.UnitsInGroup"
    GroupLevel =0
    Expression ="SchGrp.OrigEquipCost"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.LeaseTermPrimary"
    GroupLevel =0
    Expression ="SchGrp.Assignee_Bank"
    GroupLevel =0
    Expression ="Banks.BankName"
    GroupLevel =0
    Expression ="SchGrp.Assigned_Indicator"
    GroupLevel =0
    Expression ="SchGrp.AssignmentDate"
    GroupLevel =0
    Expression ="SchGrp.FundDate"
    GroupLevel =0
    Expression ="SchGrp.SalePrice"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLSch"
    GroupLevel =0
    Expression ="SchGrp.LeasePaidAdvArr"
    GroupLevel =0
    Expression ="SchGrp.LeasePymtFreq"
    GroupLevel =0
    Expression ="SchGrp.ExchangeITGroup"
    GroupLevel =0
    Expression ="SchGrp.ExchangeITMonth"
    GroupLevel =0
    Expression ="SchGrp.ExchangeITType"
    GroupLevel =0
    Expression ="SchGrp.SGrpBonus"
    GroupLevel =0
    Expression ="SchGrp.SGrpIndex"
    GroupLevel =0
    Expression ="SchGrp.SGrpIndDt"
    GroupLevel =0
    Expression ="SchGrp.SGprIndYr"
    GroupLevel =0
    Expression ="SchGrp.IndexRate"
    GroupLevel =0
    Expression ="SchGrp.SgrpYield"
    GroupLevel =0
    Expression ="SchGrp.EarlyTermNote"
    GroupLevel =0
    Expression ="SchGrp.SalesFSL"
    GroupLevel =0
    Expression ="Units.UnitSynResid"
    GroupLevel =0
    Expression ="IIf([UnitSynResid]>0,Round([UnitSynResid]/[OrigEquipCost],3),0)"
    GroupLevel =0
    Expression ="SchGrp.Fee_Projected"
    GroupLevel =0
    Expression ="SchGrp.SGrpSalesBonus"
    GroupLevel =0
    Expression ="SchGrp.SGrpSalesIndex"
    GroupLevel =0
    Expression ="SchGrp.SGrpSalesIndDt"
    GroupLevel =0
    Expression ="SchGrp.SGprSalesIndYr"
    GroupLevel =0
    Expression ="SchGrp.SgrpSalesIndRt"
    GroupLevel =0
    Expression ="SchGrp.SgrpSalesYield"
    GroupLevel =0
    Expression ="[sgrpyield]-[IndexRate]"
    GroupLevel =0
    Expression ="CCur([unitrent])"
    GroupLevel =0
    Expression ="UnitTotalOEC.UnitLRF"
    GroupLevel =0
    Expression ="Schedule.SchFSL"
    GroupLevel =0
    Expression ="FSL.FSLYR1"
    GroupLevel =0
    Expression ="SchGrp.SaleDate"
    GroupLevel =0
    Expression ="SchGrp.SgrpTermPenalty"
    GroupLevel =0
    Expression ="SchGrp.WithHeldPymts"
    GroupLevel =0
    Expression ="SyndGroupStripsUTA_Summed.UTAPer"
    GroupLevel =0
    Expression ="SchGrp.SgrpSynRebate"
    GroupLevel =0
    Expression ="SchGrp.SgrpSynPenalty"
    GroupLevel =0
    Expression ="SyndGroupStripsUTA_Summed.GroupStrips"
    GroupLevel =0
    Expression ="SyndGroupStripsUTA_Summed.GroupStripAmount"
    GroupLevel =0
    Expression ="CCur(Nz([SchGrp].[PerDiemActual]*[unitsingroup],0))"
    GroupLevel =0
    Expression ="Switch([LeasePymtFreq]=\"M\",\"Monthly\",[LeasePymtFreq]=\"Q\",\"Quarterly\",[Le"
        "asePymtFreq]=\"A\",\"Annualy\",[LeasePymtFreq]=\"S\",\"Semi-Annuly\")"
    GroupLevel =0
    Expression ="IIf(Nz([ScheduleGroupPaymentTerm],0)=0,[LeaseTermPrimary],[ScheduleGroupPaymentT"
        "erm])"
    GroupLevel =0
    Expression ="InterimFunding_Days_Rate_Group.InterimIntRate"
    GroupLevel =0
    Expression ="InterimFunding_Days_Rate_Group.AvgInterimLoanDays"
    GroupLevel =0
    Expression ="InterimFunding_Days_Rate_Group.GroupInterimExp"
    GroupLevel =0
    Expression ="FAUserFullName()"
    GroupLevel =0
    Expression ="PerDiem_Days_Rate_Group.AvgPerDiem"
    GroupLevel =0
    Expression ="PerDiem_Days_Rate_Group.PDDays"
    GroupLevel =0
    Expression ="IIf([SchGrp].[PerDiemActual]>0,[SchGrp].[PerDiemActual]*[UnitsInGroup],IIf([PerD"
        "iemGroup]>0,[PerDiemGroup],[PerDiemProj]))"
    GroupLevel =0
    Expression ="CDbl(Round([LeasePymtForGroup]/[OrigEquipCost],7))"
    GroupLevel =0
    Expression ="IIf([AllowedMilesPeriod]='P/Year',[SGprAllow1],([SGprAllow1]/([LeaseTermPrimary]"
        "-CInt(Nz([LeaseTermNotice],0)))*12))"
    GroupLevel =0
    Expression ="CCur([OrigEquipCost]/[unitsingroup])"
    GroupLevel =0
    Expression ="SchGrp.LoanDate"
    GroupLevel =0
    Expression ="SchGrp.origEquipCost"
    GroupLevel =0
    Expression ="UnitTotalOEC.UnitTotalCost"
    GroupLevel =0
    Expression ="SchGrp.LeasePymtForGroup"
    GroupLevel =0
    Expression ="CCur([UnitRent])"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="CCur(Nz([HoldbackConversion],0))"
    GroupLevel =0
    Expression ="CCur(Nz([HoldbackResidual],0))"
    GroupLevel =0
    Expression ="CDbl(Nz([WithHeldPymts],0))"
    GroupLevel =0
    Expression ="[LeasePymtForGroup]*[WithHeldPymts]"
    GroupLevel =0
    Expression ="SchGrp.LastUpdatedDate"
    GroupLevel =0
    Expression ="SchGrp.LastUpdatedBy"
    GroupLevel =0
    Expression ="SchGrp.ScheduleGroupPaymentTerm"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLPmt"
    GroupLevel =0
    Expression ="SchGrp.SGrpRSLTPmt"
    GroupLevel =0
    Expression ="Schedule.SchShortDesc"
    GroupLevel =0
    Expression ="SchGrp.SyndMiscLabel1"
    GroupLevel =0
    Expression ="SchGrp.SyndMiscValue1"
    GroupLevel =0
    Expression ="SchGrp.SyndMiscLabel2"
    GroupLevel =0
    Expression ="SchGrp.SyndMiscValue2"
    GroupLevel =0
    Expression ="Units.UnitSRent"
    GroupLevel =0
    Expression ="SchGrp.RebateUnitProjected"
    GroupLevel =0
    Expression ="Units.UnitRent"
    GroupLevel =0
    Expression ="UnitLevelTaxes.FeeTaxable"
    GroupLevel =0
    Expression ="CCur(Nz([UnitTaxes],0))"
    GroupLevel =0
    Expression ="IIf([UnitTaxes]>0,[UnitlevelTaxes].[LState],\"-\")"
    GroupLevel =0
    Expression ="IIf([UnitTaxes]>0,[UnitlevelTaxes].[LTaxRate],0)"
    GroupLevel =0
    Expression ="SchGrp.SgrpDesc"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Units.UnitSynResid"
    GroupLevel =0
    Expression ="Schedule.SchFSL"
    GroupLevel =0
    Expression ="SchGrp.SgrpSynPenalty"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLPmt"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Schedule.SchShortDesc"
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
        dbText "Name" ="SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLSch"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="SchGrp.SalesFSL"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="FSL.FSLYR1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="49"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpTermPenalty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="51"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpSynPenalty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="52"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpWHPmt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="55"
    End
    Begin
        dbText "Name" ="SchGrp.OrigEquipCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="SchGrp.UnitsInGroup"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="30"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="SchGrp.Assignee_Bank"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="33"
    End
    Begin
        dbText "Name" ="SchGrp.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="34"
    End
    Begin
        dbText "Name" ="SchGrp.AssignmentDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="35"
    End
    Begin
        dbText "Name" ="SchGrp.FundDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="36"
    End
    Begin
        dbText "Name" ="SchGrp.SalePrice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="37"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="39"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePymtForGroup"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2295"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="SchGrp.ExchangeITGroup"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="40"
    End
    Begin
        dbText "Name" ="SchGrp.ExchangeITMonth"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="41"
    End
    Begin
        dbText "Name" ="SchGrp.ExchangeITType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="42"
    End
    Begin
        dbText "Name" ="SchGrp.IndexRate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="45"
    End
    Begin
        dbText "Name" ="SchGrp.EarlyTermNote"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="46"
    End
    Begin
        dbText "Name" ="SchGrp.RebateUnitProjected"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="47"
        dbInteger "ColumnWidth" ="2370"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.Fee_Projected"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="48"
    End
    Begin
        dbText "Name" ="SchGrp.SaleDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="50"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="38"
    End
    Begin
        dbText "Name" ="SalesRent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="54"
    End
    Begin
        dbText "Name" ="Rent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="53"
    End
    Begin
        dbText "Name" ="SyndPenalty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="57"
    End
    Begin
        dbText "Name" ="StripsAmount"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="0"
    End
    Begin
        dbText "Name" ="StripsNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="58"
    End
    Begin
        dbText "Name" ="ServiceFee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="61"
    End
    Begin
        dbText "Name" ="PayFreq"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="62"
    End
    Begin
        dbText "Name" ="AdvanceArrears"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="63"
    End
    Begin
        dbText "Name" ="FeeTerm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="64"
    End
    Begin
        dbText "Name" ="FullServiceFee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="65"
    End
    Begin
        dbText "Name" ="ServFeePV"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="66"
    End
    Begin
        dbText "Name" ="PerDiemActual"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="60"
    End
    Begin
        dbText "Name" ="FullUserName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="67"
    End
    Begin
        dbText "Name" ="Guarantor"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="Index"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="43"
    End
    Begin
        dbText "Name" ="IndexDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="44"
    End
    Begin
        dbText "Name" ="IndexYear"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yield"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bonus"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesIndex"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesBonus"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesIndexYear"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesIndexRate"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesYield"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesIndexDate"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SalesUnitResidual"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="UnitResidual"
        dbInteger "ColumnWidth" ="1650"
        dbInteger "ColumnOrder" ="11"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupNum"
        dbInteger "ColumnOrder" ="27"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientShortName"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="Schedule"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RebateUnitActual"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="56"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MPY"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="72"
    End
    Begin
        dbText "Name" ="MasterLease"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CostPerUnit"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="73"
    End
    Begin
        dbText "Name" ="LRF"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="70"
    End
    Begin
        dbText "Name" ="FeeGroup"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="74"
    End
    Begin
        dbText "Name" ="FeeUnit"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="75"
    End
    Begin
        dbText "Name" ="UnitCount"
        dbInteger "ColumnWidth" ="1455"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupID"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PVFeeRate"
        dbInteger "ColumnOrder" ="71"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.loandate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTax"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AncillaryFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WithHeldAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HBConversion"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HBResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LastUpdatedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LastUpdateddate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRSLTPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiem_Days_Rate_Group.AvgPerDiem"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FeeUnitCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimLoanDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiem_Days_Rate_Group.PDDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1022"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Spread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitRVPct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SyndGroupStripsUTA_Summed.UTAPer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitRentPayment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.UnitLRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimFunding_Days_Rate_Group.InterimIntRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InterimFundingExpense"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalOEC.UnitTotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeasePymtUnit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WithHeldNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ScheduleGroupPaymentTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchShortDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SyndMiscLabel1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SyndMiscValue1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SyndMiscLabel2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SyndMiscValue2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SynMemo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiemToDisplay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LTaxRate"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-186
    Top =79
    Right =1363
    Bottom =1030
    Left =-1
    Top =-1
    Right =1338
    Bottom =179
    Left =0
    Top =192
    ColumnsShown =543
    Begin
        Left =48
        Top =-180
        Right =192
        Bottom =61
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =-180
        Right =384
        Bottom =-36
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =405
        Top =-180
        Right =549
        Bottom =-36
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =409
        Top =-29
        Right =550
        Bottom =119
        Top =0
        Name ="FSL"
        Name =""
    End
    Begin
        Left =585
        Top =-177
        Right =750
        Bottom =121
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =785
        Top =-16
        Right =929
        Bottom =315
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =788
        Top =-159
        Right =932
        Bottom =4
        Top =0
        Name ="UnitGroupPerdiemsSyndTotals"
        Name =""
    End
    Begin
        Left =241
        Top =-31
        Right =385
        Bottom =113
        Top =0
        Name ="Clients_1"
        Name =""
    End
    Begin
        Left =983
        Top =-180
        Right =1127
        Bottom =-36
        Top =0
        Name ="Banks"
        Name =""
    End
    Begin
        Left =1175
        Top =-180
        Right =1319
        Bottom =-36
        Top =0
        Name ="UnitTotalOEC"
        Name =""
    End
    Begin
        Left =1180
        Top =-35
        Right =1324
        Bottom =109
        Top =0
        Name ="SyndGroupStripsUTA_Summed"
        Name =""
    End
    Begin
        Left =1367
        Top =-180
        Right =1511
        Bottom =-36
        Top =0
        Name ="InterimFunding_Days_Rate_Group"
        Name =""
    End
    Begin
        Left =980
        Top =-34
        Right =1124
        Bottom =110
        Top =0
        Name ="PerDiem_Days_Rate_Group"
        Name =""
    End
    Begin
        Left =1356
        Top =-25
        Right =1500
        Bottom =119
        Top =0
        Name ="UnitLevelTaxes"
        Name =""
    End
End
