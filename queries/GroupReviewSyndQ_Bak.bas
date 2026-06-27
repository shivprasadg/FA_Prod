Operation =1
Option =2
Having ="(((SchGrp.SGrpID)=forms!groupform!sgrpid) And ((Units.unitsrent)>0))"
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
    Name ="UnitLevelTaxes"
    Name ="SyndGroupStripsUTA_Summed"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="Clients.clientCompanyName"
    Alias ="ClientShortName"
    Expression ="Clients.ClientShNm"
    Alias ="MasterLease"
    Expression ="MstrLease.MLNo"
    Alias ="Guarantor"
    Expression ="Clients.clientCompanyName"
    Alias ="Schedule"
    Expression ="Schedule.SchNo"
    Alias ="GroupID"
    Expression ="SchGrp.SGrpID"
    Expression ="SchGrp.SGrpNO"
    Alias ="GroupNum"
    Expression ="SchGrp.SGrpNO"
    Alias ="GroupDesc"
    Expression ="Nz([SgrpDesc],\"Units in Group: (\" & [UnitsInGroup] & \")\")"
    Expression ="SchGrp.UnitsInGroup"
    Expression ="SchGrp.OrigEquipCost"
    Expression ="SchGrp.BLCD"
    Expression ="SchGrp.LeaseTermPrimary"
    Expression ="SchGrp.LeasePymtForGroup"
    Expression ="SchGrp.Assignee_Bank"
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
    Alias ="UnitCount"
    Expression ="Count(Units.UnitID)"
    Expression ="SchGrp.SalesFSL"
    Alias ="UnitTax"
    Expression ="CCur(Nz([UnitTaxes],0))"
    Expression ="SchGrp.RebateUnitProjected"
    Alias ="UnitResidual"
    Expression ="Units.UnitSynResid"
    Alias ="SalesUnitResidual"
    Expression ="Avg(Units.unitsalesresid)"
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
    Expression ="Schedule.SchFSL"
    Expression ="FSL.FSLYR1"
    Expression ="SchGrp.SaleDate"
    Expression ="SchGrp.SgrpTermPenalty"
    Expression ="SchGrp.SgrpSynPenalty"
    Alias ="Rent"
    Expression ="Units.UNITRENT"
    Alias ="SalesRent"
    Expression ="Units.unitsrent"
    Expression ="SchGrp.WithHeldPymts"
    Alias ="RebateUnitActual"
    Expression ="SchGrp.SgrpSynRebate"
    Alias ="SyndPenalty"
    Expression ="SchGrp.SgrpSynPenalty"
    Alias ="StripsNum"
    Expression ="SyndGroupStripsUTA_Summed.GroupStrips"
    Alias ="StripsAmount"
    Expression ="SyndGroupStripsUTA_Summed.GroupStripAmount"
    Alias ="PerDiemActual"
    Expression ="CCur(Nz([SchGrp].[PerDiemActual]*[unitsingroup]))"
    Alias ="ServiceFee"
    Expression ="SchGrp.SGrpFSLPmt"
    Alias ="AncillaryFee"
    Expression ="UnitLevelTaxes.FeeTaxable"
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
        "tTerm]))*([SGrpFSLPmt]+Nz([AncillaryFee],0)))*[Unitsingroup]"
    Alias ="PVFeeRate"
    Expression ="CDbl(Nz([forms]![groupform]![pvfeerate],0.05))"
    Alias ="ServFeePV"
    Expression ="IIf([forms]![groupform]![PVFeesYN]=True And [forms]![groupform]![ServFeePassTrou"
        "ghYN]=False,Round(PV(CDbl((Nz([PVFeeRate],0.05)/12)),[FeeTerm],([SGrpFSLPmt]*[Un"
        "itsInGroup])*-1,0,1),2),0)"
    Alias ="IntrimIntRate"
    Expression ="forms.groupform![InterimLoanRate]"
    Alias ="IntrimLoanDays"
    Expression ="DateDiff(\"d\",[LoanDate],[FundDate])"
    Alias ="IntrimFundingExpense"
    Expression ="IIf([forms].[groupform]![CalcIntrimYN]=True,CDbl(IIf(IsDate([LoanDate]),[OrigEqu"
        "ipCost]+FV(([IntrimIntRate]/360),[IntrimLoanDays],0,[OrigEquipCost],1),0)),0)"
    Alias ="FullUserName"
    Expression ="FAUserFullName()"
    Alias ="PerdiemGroup"
    Expression ="IIf([forms].[groupform]![CalcPerDiemYN]=True,Max(CDbl(Nz([PerdiemTotal],0))),CCu"
        "r(Nz([SchGrp].[PerDiemActual]*[unitsingroup])))"
    Alias ="LRF"
    Expression ="CDbl(Round([LeasePymtForGroup]/[OrigEquipCost],7))"
    Alias ="MPY"
    Expression ="IIf([AllowedMilesPeriod]='P/TERM',[SGprAllow1]/([LeaseTermPrimary]-Nz([LeaseTerm"
        "Notice],0)),[SGprAllow1])"
    Alias ="CostPerUnit"
    Expression ="CCur([OrigEquipCost]/[unitsingroup])"
    Alias ="FeeGroup"
    Expression ="IIf([SalePrice]>0,([SalePrice]-[OrigEquipCost]),0)"
    Alias ="FeeUnit"
    Expression ="CCur(IIf([SalePrice]>0,(([SalePrice]-[OrigEquipCost]))/[UnitsInGroup],0))"
    Expression ="SchGrp.LoanDate"
    Alias ="HBConversion"
    Expression ="CCur(Nz([HoldbackConversion],0))"
    Alias ="HBResidual"
    Expression ="CCur(Nz([HoldbackResidual],0))"
    Alias ="WithHeldAmount"
    Expression ="[LeasePymtForGroup]*[WithHeldPymts]"
    Alias ="LocState"
    Expression ="IIf([UnitTaxes]>0,[UnitlevelTaxes].[LocState],\"-\")"
    Alias ="LTaxRate"
    Expression ="IIf([UnitTaxes]>0,[UnitlevelTaxes].[LTaxRate],0)"
    Expression ="SchGrp.LastUpdatedBy"
    Expression ="SchGrp.LastUpdatedDate"
    Expression ="SchGrp.SGrpRSLTPmt"
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
    LeftTable ="Units"
    RightTable ="UnitLevelTaxes"
    Expression ="Units.UnitID = UnitLevelTaxes.UnitID"
    Flag =1
    LeftTable ="SchGrp"
    RightTable ="SyndGroupStripsUTA_Summed"
    Expression ="SchGrp.SGrpID = SyndGroupStripsUTA_Summed.GroupID"
    Flag =1
End
Begin Groups
    Expression ="Clients.ClientShNm"
    GroupLevel =0
    Expression ="MstrLease.MLNo"
    GroupLevel =0
    Expression ="Clients.clientCompanyName"
    GroupLevel =0
    Expression ="Schedule.SchNo"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
    Expression ="SchGrp.SGrpNO"
    GroupLevel =0
    Expression ="Nz([SgrpDesc],\"Units in Group: (\" & [UnitsInGroup] & \")\")"
    GroupLevel =0
    Expression ="SchGrp.UnitsInGroup"
    GroupLevel =0
    Expression ="SchGrp.OrigEquipCost"
    GroupLevel =0
    Expression ="SchGrp.BLCD"
    GroupLevel =0
    Expression ="SchGrp.LeaseTermPrimary"
    GroupLevel =0
    Expression ="SchGrp.LeasePymtForGroup"
    GroupLevel =0
    Expression ="SchGrp.Assignee_Bank"
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
    Expression ="CCur(Nz([UnitTaxes],0))"
    GroupLevel =0
    Expression ="SchGrp.RebateUnitProjected"
    GroupLevel =0
    Expression ="Units.UnitSynResid"
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
    Expression ="Schedule.SchFSL"
    GroupLevel =0
    Expression ="FSL.FSLYR1"
    GroupLevel =0
    Expression ="SchGrp.SaleDate"
    GroupLevel =0
    Expression ="SchGrp.SgrpTermPenalty"
    GroupLevel =0
    Expression ="Units.UNITRENT"
    GroupLevel =0
    Expression ="Units.unitsrent"
    GroupLevel =0
    Expression ="SchGrp.WithHeldPymts"
    GroupLevel =0
    Expression ="SchGrp.SgrpSynRebate"
    GroupLevel =0
    Expression ="SchGrp.SgrpSynPenalty"
    GroupLevel =0
    Expression ="SyndGroupStripsUTA_Summed.GroupStrips"
    GroupLevel =0
    Expression ="SyndGroupStripsUTA_Summed.GroupStripAmount"
    GroupLevel =0
    Expression ="CCur(Nz([SchGrp].[PerDiemActual]*[unitsingroup]))"
    GroupLevel =0
    Expression ="SchGrp.SGrpFSLPmt"
    GroupLevel =0
    Expression ="UnitLevelTaxes.FeeTaxable"
    GroupLevel =0
    Expression ="Switch([LeasePymtFreq]=\"M\",\"Monthly\",[LeasePymtFreq]=\"Q\",\"Quarterly\",[Le"
        "asePymtFreq]=\"A\",\"Annualy\",[LeasePymtFreq]=\"S\",\"Semi-Annuly\")"
    GroupLevel =0
    Expression ="IIf(Nz([ScheduleGroupPaymentTerm],0)=0,[LeaseTermPrimary],[ScheduleGroupPaymentT"
        "erm])"
    GroupLevel =0
    Expression ="forms.groupform![InterimLoanRate]"
    GroupLevel =0
    Expression ="DateDiff(\"d\",[LoanDate],[FundDate])"
    GroupLevel =0
    Expression ="FAUserFullName()"
    GroupLevel =0
    Expression ="CDbl(Round([LeasePymtForGroup]/[OrigEquipCost],7))"
    GroupLevel =0
    Expression ="IIf([AllowedMilesPeriod]='P/TERM',[SGprAllow1]/([LeaseTermPrimary]-Nz([LeaseTerm"
        "Notice],0)),[SGprAllow1])"
    GroupLevel =0
    Expression ="CCur([OrigEquipCost]/[unitsingroup])"
    GroupLevel =0
    Expression ="SchGrp.LoanDate"
    GroupLevel =0
    Expression ="CCur(Nz([HoldbackConversion],0))"
    GroupLevel =0
    Expression ="CCur(Nz([HoldbackResidual],0))"
    GroupLevel =0
    Expression ="[LeasePymtForGroup]*[WithHeldPymts]"
    GroupLevel =0
    Expression ="IIf([UnitTaxes]>0,[UnitlevelTaxes].[LocState],\"-\")"
    GroupLevel =0
    Expression ="IIf([UnitTaxes]>0,[UnitlevelTaxes].[LTaxRate],0)"
    GroupLevel =0
    Expression ="SchGrp.LastUpdatedBy"
    GroupLevel =0
    Expression ="SchGrp.LastUpdatedDate"
    GroupLevel =0
    Expression ="SchGrp.SGrpRSLTPmt"
    GroupLevel =0
    Expression ="([unitrent]+IIf([schfsl]=No,0,[sgrpfslpmt]-[fslyr1]))*[UnitsInGroup]"
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
        dbInteger "ColumnOrder" ="1"
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
        dbText "Name" ="IntrimFundingExpense"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2520"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="68"
    End
    Begin
        dbText "Name" ="PerdiemGroup"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="69"
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
        dbText "Name" ="IntrimIntRate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="IntrimLoanDays"
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
End
Begin
    State =0
    Left =0
    Top =40
    Right =1549
    Bottom =991
    Left =-1
    Top =-1
    Right =1517
    Bottom =76
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =253
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="MstrLease"
        Name =""
    End
    Begin
        Left =405
        Top =12
        Right =549
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =409
        Top =163
        Right =550
        Bottom =311
        Top =0
        Name ="FSL"
        Name =""
    End
    Begin
        Left =585
        Top =15
        Right =750
        Bottom =313
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =785
        Top =176
        Right =929
        Bottom =507
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =791
        Top =69
        Right =935
        Bottom =232
        Top =0
        Name ="UnitGroupPerdiemsSyndTotals"
        Name =""
    End
    Begin
        Left =241
        Top =161
        Right =385
        Bottom =305
        Top =0
        Name ="Clients_1"
        Name =""
    End
    Begin
        Left =969
        Top =67
        Right =1113
        Bottom =211
        Top =0
        Name ="UnitLevelTaxes"
        Name =""
    End
    Begin
        Left =1129
        Top =13
        Right =1273
        Bottom =157
        Top =0
        Name ="SyndGroupStripsUTA_Summed"
        Name =""
    End
End
