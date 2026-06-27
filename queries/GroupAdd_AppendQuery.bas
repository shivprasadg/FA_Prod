Operation =3
Name ="schgrp"
Option =0
Begin InputTables
    Name ="GroupFormAdd_WorkTable"
    Alias ="W"
End
Begin OutputColumns
    Name ="SGrpSchID"
    Expression ="W.SGrpSchID"
    Name ="SGrpNO"
    Expression ="W.SGrpNO"
    Name ="GroupAddedDate"
    Expression ="W.GroupAddedDate"
    Name ="UnitsInGroup"
    Expression ="W.UnitsInGroup"
    Name ="AssigneeBankID"
    Expression ="W.AssigneeBankID"
    Name ="Assignee_Bank"
    Expression ="W.Assignee_Bank"
    Name ="Assigned_Indicator"
    Expression ="W.Assigned_Indicator"
    Name ="AssignmentDate"
    Expression ="W.AssignmentDate"
    Name ="BPChgLO"
    Expression ="W.BPChgLO"
    Name ="BPChgLO_Up"
    Expression ="W.BPChgLO_Up"
    Name ="BPChgLO_Down"
    Expression ="W.BPChgLO_Down"
    Name ="BPChgSynd"
    Expression ="W.BPChgSynd"
    Name ="BPChgSynd_Up"
    Expression ="W.BPChgSynd_Up"
    Name ="BPChgSynd_Down"
    Expression ="W.BPChgSynd_Down"
    Name ="BLCD"
    Expression ="W.BLCD"
    Name ="BLCD_Forecast"
    Expression ="W.BLCD_Forecast"
    Name ="BLCD_Projected"
    Expression ="W.BLCD_Projected"
    Name ="BrokerFee"
    Expression ="W.BrokerFee"
    Name ="BuyOutPrice"
    Expression ="W.BuyOutPrice"
    Name ="BuyOutNote"
    Expression ="W.BuyOutNote"
    Name ="BuyOutTBD"
    Expression ="W.BuyOutTBD"
    Name ="CapLeaseBuyOut"
    Expression ="W.CapLeaseBuyOut"
    Name ="EarlyTermOptional"
    Expression ="W.EarlyTermOptional"
    Name ="EarlyTermNote"
    Expression ="W.EarlyTermNote"
    Name ="ExchangeITGroup"
    Expression ="W.ExchangeITGroup"
    Name ="ExchangeITMonth"
    Expression ="W.ExchangeITMonth"
    Name ="ExchangeITType"
    Expression ="W.ExchangeITType"
    Name ="Fee_Projected"
    Expression ="W.Fee_Projected"
    Name ="FMVBuyOutOptional"
    Expression ="W.FMVBuyOutOptional"
    Name ="FundDate"
    Expression ="W.FundDate"
    Name ="FundDate_Projected"
    Expression ="W.FundDate_Projected"
    Name ="LeaseType"
    Expression ="W.LeaseType"
    Name ="LeaseTermPrimary"
    Expression ="W.LeaseTermPrimary"
    Name ="LeasePymtForGroup"
    Expression ="W.LeasePymtForGroup"
    Name ="LeasePymtFreq"
    Expression ="W.LeasePymtFreq"
    Name ="LeasePaidAdvArr"
    Expression ="W.LeasePaidAdvArr"
    Name ="LeaseTermNotice"
    Expression ="W.LeaseTermNotice"
    Name ="LRFGroup"
    Expression ="W.LRFGroup"
    Name ="LoanBank"
    Expression ="W.LoanBank"
    Name ="LoanDate"
    Expression ="W.LoanDate"
    Name ="LoanAmount"
    Expression ="W.LoanAmount"
    Name ="LoanType"
    Expression ="W.LoanType"
    Name ="OrigEquipCost"
    Expression ="W.OrigEquipCost"
    Name ="SaleDate"
    Expression ="W.SaleDate"
    Name ="SynFee"
    Expression ="W.SynFee"
    Name ="SalePrice"
    Expression ="W.SalePrice"
    Name ="Strips_Projected"
    Expression ="W.Strips_Projected"
    Name ="RebateUnitProjected"
    Expression ="W.RebateUnitProjected"
    Name ="SGrpRtTBD"
    Expression ="W.SGrpRtTBD"
    Name ="LeaseExtOption1Term"
    Expression ="W.LeaseExtOption1Term"
    Name ="LeaseExtOption1LRF"
    Expression ="W.LeaseExtOption1LRF"
    Name ="LeaseExtOption1Rent"
    Expression ="W.LeaseExtOption1Rent"
    Name ="BillToAddressID"
    Expression ="W.BillToAddressID"
    Name ="SGrpOpt2"
    Expression ="W.SGrpOpt2"
    Name ="SGrpO2LRF"
    Expression ="W.SGrpO2LRF"
    Name ="SGrpO2uRent"
    Expression ="W.SGrpO2uRent"
    Name ="SGrpOpt3"
    Expression ="W.SGrpOpt3"
    Name ="SGrpO3LRF"
    Expression ="W.SGrpO3LRF"
    Name ="SGrpO3uRent"
    Expression ="W.SGrpO3uRent"
    Name ="SGrpFABill"
    Expression ="W.SGrpFABill"
    Name ="SGrpExpDt"
    Expression ="W.SGrpExpDt"
    Name ="SgrpPayoffDt"
    Expression ="W.SgrpPayoffDt"
    Name ="SGrpEngReb"
    Expression ="W.SGrpEngReb"
    Name ="SgrpEngMi"
    Expression ="W.SgrpEngMi"
    Name ="SGprAllow1"
    Expression ="W.SGprAllow1"
    Name ="SGprAllow2"
    Expression ="W.SGprAllow2"
    Name ="SGprAllow3"
    Expression ="W.SGprAllow3"
    Name ="SGprAllow4"
    Expression ="W.SGprAllow4"
    Name ="SGprAllow5"
    Expression ="W.SGprAllow5"
    Name ="SGprAllow6"
    Expression ="W.SGprAllow6"
    Name ="SGrpMH1"
    Expression ="W.SGrpMH1"
    Name ="SGrpMH2"
    Expression ="W.SGrpMH2"
    Name ="SGrpMH3"
    Expression ="W.SGrpMH3"
    Name ="SGrpMH4"
    Expression ="W.SGrpMH4"
    Name ="SGrpMH5"
    Expression ="W.SGrpMH5"
    Name ="SGrpMH6"
    Expression ="W.SGrpMH6"
    Name ="SGrpov1"
    Expression ="W.SGrpov1"
    Name ="SGrpov2"
    Expression ="W.SGrpov2"
    Name ="SGrpov3"
    Expression ="W.SGrpov3"
    Name ="SGrpov4"
    Expression ="W.SGrpov4"
    Name ="SGrpov5"
    Expression ="W.SGrpov5"
    Name ="SGrpov6"
    Expression ="W.SGrpov6"
    Name ="SgrpEq1"
    Expression ="W.SgrpEq1"
    Name ="SgrpEq2"
    Expression ="W.SgrpEq2"
    Name ="SgrpEq3"
    Expression ="W.SgrpEq3"
    Name ="SgrpEq4"
    Expression ="W.SgrpEq4"
    Name ="SgrpEq5"
    Expression ="W.SgrpEq5"
    Name ="SgrpEq6"
    Expression ="W.SgrpEq6"
    Name ="AllowedMilesPer"
    Expression ="W.AllowedMilesPer"
    Name ="AllowedMilesPeriod"
    Expression ="W.AllowedMilesPeriod"
    Name ="SGSL"
    Expression ="W.SGSL"
    Name ="SGrpRefR"
    Expression ="W.SGrpRefR"
    Name ="SGrpRHr"
    Expression ="W.SGrpRHr"
    Name ="SGrpRALO"
    Expression ="W.SGrpRALO"
    Name ="SGrpRALB"
    Expression ="W.SGrpRALB"
    Name ="SGrpRALSign"
    Expression ="W.SGrpRALSign"
    Name ="SGRPNoPD"
    Expression ="W.SGRPNoPD"
    Name ="SGRPINTPD"
    Expression ="W.SGRPINTPD"
    Name ="SGRPINTRT"
    Expression ="W.SGRPINTRT"
    Name ="SGrpNoSlv"
    Expression ="W.SGrpNoSlv"
    Name ="SGrpFSLPmt"
    Expression ="W.SGrpFSLPmt"
    Name ="PerDiemActual"
    Expression ="W.PerDiemActual"
    Name ="PerDiemProjected"
    Expression ="W.PerDiemProjected"
    Name ="PerDiemStartsOn"
    Expression ="W.PerDiemStartsOn"
    Name ="SGrpFSLSch"
    Expression ="W.SGrpFSLSch"
    Name ="SGrpSalesBonus"
    Expression ="W.SGrpSalesBonus"
    Name ="SGrpSalesIndex"
    Expression ="W.SGrpSalesIndex"
    Name ="SGrpSalesIndDt"
    Expression ="W.SGrpSalesIndDt"
    Name ="SGprSalesIndYr"
    Expression ="W.SGprSalesIndYr"
    Name ="SgrpSalesIndRt"
    Expression ="W.SgrpSalesIndRt"
    Name ="SgrpSalesYield"
    Expression ="W.SgrpSalesYield"
    Name ="SgrpSalesSpread"
    Expression ="W.SgrpSalesSpread"
    Name ="SGrpBonus"
    Expression ="W.SGrpBonus"
    Name ="SGrpIndex"
    Expression ="W.SGrpIndex"
    Name ="SGrpIndDt"
    Expression ="W.SGrpIndDt"
    Name ="SGprIndYr"
    Expression ="W.SGprIndYr"
    Name ="IndexRate"
    Expression ="W.IndexRate"
    Name ="SgrpYield"
    Expression ="W.SgrpYield"
    Name ="SgrpSpread"
    Expression ="W.SgrpSpread"
    Name ="SGrpFSLTax"
    Expression ="W.SGrpFSLTax"
    Name ="SGRPPLB"
    Expression ="W.SGRPPLB"
    Name ="SgrpBnkNo"
    Expression ="W.SgrpBnkNo"
    Name ="SGrpFAID"
    Expression ="W.SGrpFAID"
    Name ="POVendor"
    Expression ="W.POVendor"
    Name ="POSentDt"
    Expression ="W.POSentDt"
    Name ="PO2Vendor"
    Expression ="W.PO2Vendor"
    Name ="PO2SentDt"
    Expression ="W.PO2SentDt"
    Name ="PO3Vendor"
    Expression ="W.PO3Vendor"
    Name ="PO3SentDt"
    Expression ="W.PO3SentDt"
    Name ="PO4Vendor"
    Expression ="W.PO4Vendor"
    Name ="PO4SentDt"
    Expression ="W.PO4SentDt"
    Name ="PO5Vendor"
    Expression ="W.PO5Vendor"
    Name ="PO5SentDt"
    Expression ="W.PO5SentDt"
    Name ="PO6Vendor"
    Expression ="W.PO6Vendor"
    Name ="PO6SentDt"
    Expression ="W.PO6SentDt"
    Name ="M-T-MExt"
    Expression ="W.[M-T-MExt]"
    Name ="M-T-MExt1"
    Expression ="W.[M-T-MExt1]"
    Name ="M-T-MExt2"
    Expression ="W.[M-T-MExt2]"
    Name ="M-T-MExt3"
    Expression ="W.[M-T-MExt3]"
    Name ="FAFee"
    Expression ="W.FAFee"
    Name ="SalesFSL"
    Expression ="W.SalesFSL"
    Name ="CustPO1"
    Expression ="W.CustPO1"
    Name ="CustPO2"
    Expression ="W.CustPO2"
    Name ="CustPO3"
    Expression ="W.CustPO3"
    Name ="CustPO4"
    Expression ="W.CustPO4"
    Name ="SGrpIntBnk"
    Expression ="W.SGrpIntBnk"
    Name ="SgrpCommInd"
    Expression ="W.SgrpCommInd"
    Name ="AllowedHoursPer"
    Expression ="W.AllowedHoursPer"
    Name ="AllowedHoursPeriod"
    Expression ="W.AllowedHoursPeriod"
    Name ="AllowedMilesNote"
    Expression ="W.AllowedMilesNote"
    Name ="SgrpTermPenalty"
    Expression ="W.SgrpTermPenalty"
    Name ="SGrpExtFSLPmt"
    Expression ="W.SGrpExtFSLPmt"
    Name ="ExtendFSLPymt"
    Expression ="W.ExtendFSLPymt"
    Name ="SGrpRSLTPmt"
    Expression ="W.SGrpRSLTPmt"
    Name ="SgrpFA4Assign"
    Expression ="W.SgrpFA4Assign"
    Name ="SgrpCustPmt"
    Expression ="W.SgrpCustPmt"
    Name ="SgrpSynPenalty"
    Expression ="W.SgrpSynPenalty"
    Name ="SgrpSynRebate"
    Expression ="W.SgrpSynRebate"
    Name ="SgrpDesc"
    Expression ="W.SgrpDesc"
    Name ="WithHeldPymts"
    Expression ="W.WithHeldPymts"
    Name ="SynMemo"
    Expression ="W.SynMemo"
    Name ="MTMRate"
    Expression ="W.MTMRate"
    Name ="MTMRent"
    Expression ="W.MTMRent"
    Name ="PONum1"
    Expression ="W.PONum1"
    Name ="PONum2"
    Expression ="W.PONum2"
    Name ="PONum3"
    Expression ="W.PONum3"
    Name ="PONum4"
    Expression ="W.PONum4"
    Name ="ScheduleGroupPaymentTerm"
    Expression ="W.ScheduleGroupPaymentTerm"
    Name ="LoanSecurityAgmtYN"
    Expression ="W.LoanSecurityAgmtYN"
    Name ="OmnibusAgmtYN"
    Expression ="W.OmnibusAgmtYN"
    Name ="ProgressPymtAgmtYN"
    Expression ="W.ProgressPymtAgmtYN"
    Name ="ThirdPartyPayAssgnmtYN"
    Expression ="W.ThirdPartyPayAssgnmtYN"
    Name ="CustomDocumentNameYN"
    Expression ="W.CustomDocumentNameYN"
    Name ="CustomDocumentName"
    Expression ="W.CustomDocumentName"
    Name ="BLCDSetBy"
    Expression ="W.BLCDSetBy"
    Name ="HoldbackResidual"
    Expression ="W.HoldbackResidual"
    Name ="HoldbackConversion"
    Expression ="W.HoldbackConversion"
    Name ="FAServicer"
    Expression ="W.FAServicer"
    Name ="FAInvoicing"
    Expression ="W.FAInvoicing"
    Name ="FACollecting"
    Expression ="W.FACollecting"
    Name ="FA3rdPartyPayingAgent"
    Expression ="W.FA3rdPartyPayingAgent"
    Name ="LPAFCost"
    Expression ="W.LPAFCost"
    Name ="SyndMiscLabel1"
    Expression ="W.SyndMiscLabel1"
    Name ="SyndMiscValue1"
    Expression ="W.SyndMiscValue1"
    Name ="SyndMiscLabel2"
    Expression ="W.SyndMiscLabel2"
    Name ="SyndMiscValue2"
    Expression ="W.SyndMiscValue2"
    Name ="PVRate"
    Expression ="W.PVRate"
    Name ="InterimLoanRate"
    Expression ="W.InterimLoanRate"
    Name ="ResidualGuaranteeDocYN"
    Expression ="W.ResidualGuaranteeDocYN"
    Name ="RGDCasualtyYN"
    Expression ="W.RGDCasualtyYN"
    Name ="RGDExchangeITYN"
    Expression ="W.RGDExchangeITYN"
    Name ="RGDRepurchaseTN"
    Expression ="W.RGDRepurchaseTN"
    Name ="RGDTerminationYN"
    Expression ="W.RGDTerminationYN"
    Name ="ShortFallAgreementYN"
    Expression ="W.ShortFallAgreementYN"
    Name ="SFACasualtyYN"
    Expression ="W.SFACasualtyYN"
    Name ="SFAExchangeITYN"
    Expression ="W.SFAExchangeITYN"
    Name ="SFARepurchaseTN"
    Expression ="W.SFARepurchaseTN"
    Name ="SFATerminationYN"
    Expression ="W.SFATerminationYN"
    Name ="IndemnificationDocYN"
    Expression ="W.IndemnificationDocYN"
    Name ="IDFMileageYN"
    Expression ="W.IDFMileageYN"
    Name ="IDFLateFeesYN"
    Expression ="W.IDFLateFeesYN"
    Name ="LesseeDefaults"
    Expression ="W.LesseeDefaults"
    Name ="LesseeObligations"
    Expression ="W.LesseeObligations"
    Name ="LesseeDefaultsAll"
    Expression ="W.LesseeDefaultsAll"
    Name ="LesseeDefaultsCertain"
    Expression ="W.LesseeDefaultsCertain"
    Name ="LesseeObligationsAll"
    Expression ="W.LesseeObligationsAll"
    Name ="LesseeObligationsCertain"
    Expression ="W.LesseeObligationsCertain"
    Name ="UseSpecialRemitGrp"
    Expression ="W.UseSpecialRemitGrp"
    Name ="SpecialRemitAddrGrp"
    Expression ="W.SpecialRemitAddrGrp"
    Name ="ATFGroupSpecID"
    Expression ="W.ATFGroupSpecID"
    Name ="ATFID"
    Expression ="W.ATFID"
    Name ="BillingInstructions"
    Expression ="W.BillingInstructions"
    Name ="TaxInstructions"
    Expression ="W.TaxInstructions"
    Name ="PVFeesYN"
    Expression ="W.PVFeesYN"
    Name ="ServFeePassTroughYN"
    Expression ="W.ServFeePassTroughYN"
    Name ="POQuoteNums"
    Expression ="W.POQuoteNums"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.Strips_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.UnitsInGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LeaseTermNotice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.IndemnificationDocYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BPChgSynd_Down"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpCommInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.POQuoteNums"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PO3Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpOpt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.[M-T-MExt2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.AssigneeBankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LastUpdatedValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpO3uRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BPChgLO_Down"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGprAllow5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LeaseExtOption1Term"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpov4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpWHPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LoanDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PVRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BPChgLO_Down"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BuyOutTBD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpRtTBD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.GroupAddedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BPChgSynd_Up"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LoanType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.CustPO3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BPChgLO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BuyOutNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BPChgLO_Up"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BPChgLO_Up"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.ExchangeITType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BPChgSynd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BLCD_Forecast"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.[M-T-MExt3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpRALO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpO2LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.CapLeaseBuyOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.POVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.EarlyTermOptional"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BrokerFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BuyOutPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.EarlyTermNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpO3LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.EarlyTermNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpEq2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.ExchangeITGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.IDFMileageYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.RebateUnitProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LRFGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.ExchangeITMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpEq3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpOpt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.RGDTerminationYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGRPPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.ExchangeITGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpUT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.Fee_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.IDFMileageYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PONum2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.FMVBuyOutOptional"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BillingInstructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.FundDate_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SynFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LeaseType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpPayoffDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpPayoffDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpFSLSch"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpMH2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LeaseExtOption1Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SFATerminationYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.MaintTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PO4SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PO2Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LeasePymtForGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpBonus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LoanAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.OrigEquipCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.HoursOverPct1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.FACollecting"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PONum3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SaleDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpO2uRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.CustPO2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LastUpdatedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LeaseExtOption1LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpRtTBD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SaleDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BillToAddressID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpSalesSpread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpFABill"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpExpDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PO5Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpEngReb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpEngMi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGprAllow1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.[SGPRHU/A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGprAllow2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpFSLInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGprAllow3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGprAllow4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGprAllow6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpMH1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.FAServicer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.CustPO2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpFSLTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpMH3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpMH4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpMH5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpMH6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpov1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.MilesOverPct2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpov2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PONum4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.MilesOver1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpov3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.ScheduleGroupPaymentTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpNoSlv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpov4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpov5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpov5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpov6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpEq1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpExpDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpEq2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpEq4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpEq5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.sgrphrnt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpEq6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.[SGPRU/A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpRefR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpSalesIndDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpRHr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpRALO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LRFGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpRALB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.MTMRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpFA4Assign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpRALSign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.Strips_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGRPNoPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpSalesBonus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGRPINTPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGRPINTRT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PerDiemActual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PerDiemProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PerDiemStartsOn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.RGDRepurchaseTN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpSalesBonus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.CustPO1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PO3SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpSalesIndex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.RebateUnitProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.HoursOver3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGprSalesIndYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpSalesIndRt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.HoursOver2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpSalesYield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpIndex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpEq4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpIndDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BuyOutPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGprIndYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.sgrpmint"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.MaintProv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PO4Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.IndexRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpYield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.FA3rdPartyPayingAgent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.HoldbackResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.FAFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpSpread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.ThirdPartyPayAssgnmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpFSLInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PONum3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpFSLRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LesseeObligations"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PO2Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGRPPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpBnkNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SalesFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.POVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.POSentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGprAllow3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.MilesOver2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.[M-T-MExt1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PO2SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PO3Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpIntBnk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PO3SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.[M-T-MExt]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PO4SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PO5Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SFACasualtyYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.Fee_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PO5SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpEq6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.HoursOverPct1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PO6Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PO6SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.UnitPctOver"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.[M-T-MExt2]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BPChgSynd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.[M-T-MExt3]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.MilesOver2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.MaintTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SalesFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.CustPO1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.UnitsInGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpCustPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.CustPO3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.CustPO4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpIntBnk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.[SGPRHU/A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGRPINTRT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpHUT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.sgrpmint"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpTermPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.ServFeePassTroughYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.MaintProv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpExtFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpExTxFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpFA4Assign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SGrpRSLTPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PONum1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpSynPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.ShortFallAgreementYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpSynRebate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.ProgressPymtAgmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.sgrphrnt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SgrpDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SynMemo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.MTMRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpRALSign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.MTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.UnitPctOver"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.MilesOver1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.MilesOver3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.MilesOverPct1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.OrigEquipCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.MilesOverPct2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpEq3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BrokerFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.MilesOverPct3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.HoursOver1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.HoursOver2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.HoursOverPct2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.HoursOverPct3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PONum4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpUT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.ScheduleGroupPaymentTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LoanSecurityAgmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.OmnibusAgmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.ProgressPymtAgmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.CustomDocumentNameYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.CustomDocumentName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BLCDSetBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpFSLTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.HoldbackResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.HoldbackConversion"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.FAServicer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.FAInvoicing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpSynRebate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.FA3rdPartyPayingAgent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LPAFCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpov3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.InterimLoanRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SyndMiscLabel1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.[M-T-MExt1]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.ResidualGuaranteeDocYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SyndMiscValue1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LeaseExtOption1LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SyndMiscLabel2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.MTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SyndMiscValue2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.RGDCasualtyYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.RGDExchangeITYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.RGDTerminationYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SFACasualtyYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.EarlyTermOptional"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SFAExchangeITYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGRPINTPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SFARepurchaseTN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SFATerminationYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.UseSpecialRemitGrp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpFABill"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.IDFLateFeesYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpOpt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LesseeDefaults"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LesseeObligations"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LesseeDefaultsAll"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LesseeDefaultsCertain"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LesseeObligationsAll"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpov1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LesseeObligationsCertain"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpIndDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LastUpdatedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LastUpdatedField"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PerDiemStartsOn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.LastUpdatedValue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.UseSpecialRemitGrp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.SpecialRemitAddrGrp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.BillingInstructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.TaxInstructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpO3LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.PVFeesYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LoanType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFormAdd_WorkTable.ServFeePassTroughYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.IndexRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.GroupAddedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.AssigneeBankID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SyndMiscValue2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BPChgLO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BPChgSynd_Up"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BPChgSynd_Down"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpBonus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BLCD_Forecast"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BuyOutNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BuyOutTBD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.CapLeaseBuyOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpSalesSpread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.ExchangeITMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.ExchangeITType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.FMVBuyOutOptional"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpCustPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.FundDate_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LeaseType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LeasePymtForGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LPAFCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.FAFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGprAllow1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LeaseTermNotice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LoanDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LoanAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SynFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpCommInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LeaseExtOption1Term"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LeaseExtOption1Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BillToAddressID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpO2LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpO2uRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpOpt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LesseeDefaultsCertain"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpO3uRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpEngReb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpEngMi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGprAllow2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGprAllow4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGprAllow5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpFSLRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGprAllow6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.InterimLoanRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpSalesIndex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpMH1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpMH2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpMH3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpMH4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpExTxFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpMH5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpMH6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpov2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpov6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SFARepurchaseTN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpEq1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LastUpdatedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpEq5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.OmnibusAgmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.[SGPRU/A]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpRefR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpRHr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpRALB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGRPNoPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpNoSlv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.ResidualGuaranteeDocYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PerDiemActual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PerDiemProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpFSLSch"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpSalesIndDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGprSalesIndYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpSalesIndRt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpSalesYield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpIndex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGprIndYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpYield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpSpread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LesseeObligationsCertain"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpBnkNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.POSentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PO2SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PO6SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PO4Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PO5SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PO6Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.[M-T-MExt]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.CustPO4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.RGDExchangeITYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpHUT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpTermPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.FACollecting"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpExtFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SGrpRSLTPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpSynPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SgrpWHPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.HoursOverPct3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SynMemo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.MilesOver3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.MilesOverPct1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.MilesOverPct3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.HoursOver1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.HoursOver3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.HoursOverPct2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PONum1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PONum2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LoanSecurityAgmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.ThirdPartyPayAssgnmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.CustomDocumentNameYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SyndMiscValue1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.CustomDocumentName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.BLCDSetBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.HoldbackConversion"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.FAInvoicing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LesseeDefaults"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SyndMiscLabel1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SyndMiscLabel2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PVRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.RGDCasualtyYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.RGDRepurchaseTN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.ShortFallAgreementYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SFAExchangeITYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.IndemnificationDocYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.IDFLateFeesYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LesseeDefaultsAll"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LesseeObligationsAll"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LastUpdatedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.POQuoteNums"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.LastUpdatedField"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.SpecialRemitAddrGrp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.TaxInstructions"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="W.PVFeesYN"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =18
    Top =79
    Right =1241
    Bottom =1125
    Left =-1
    Top =-1
    Right =1199
    Bottom =813
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="W"
        Name =""
    End
End
