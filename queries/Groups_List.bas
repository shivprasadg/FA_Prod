Operation =1
Option =0
Begin InputTables
    Name ="Clients"
    Name ="MstrLease"
    Name ="Schedule"
    Name ="SchGrp"
End
Begin OutputColumns
    Expression ="Clients.ClientShNm"
    Expression ="Clients.clientID"
    Expression ="MstrLease.MLNo"
    Expression ="Schedule.SchID"
    Expression ="Schedule.SchNo"
    Alias ="GroupID"
    Expression ="SchGrp.SGrpID"
    Expression ="SchGrp.*"
    Expression ="Schedule.SchBnkClNO"
    Expression ="Schedule.BLCDACDT"
    Expression ="MstrLease.MLOrig"
    Expression ="MstrLease.MLDt"
    Expression ="Clients.ClientGroupID"
    Expression ="MstrLease.MLOrig"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="MstrLease"
    Expression ="Clients.clientID = MstrLease.Client"
    Flag =1
    LeftTable ="MstrLease"
    RightTable ="Schedule"
    Expression ="MstrLease.MLID = Schedule.SchMLRef"
    Flag =1
    LeftTable ="Schedule"
    RightTable ="SchGrp"
    Expression ="Schedule.SchID = SchGrp.SGrpSchID"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.ClientShNm"
    Flag =0
    Expression ="MstrLease.MLNo"
    Flag =0
    Expression ="Schedule.SchID"
    Flag =0
    Expression ="SchGrp.SGrpNO"
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
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRtTBD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpOpt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpO2LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpO2uRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpOpt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpO3LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpO3uRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFABill"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpExpDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpPayoffDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpEngReb"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpEngMi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprAllow1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprAllow2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprAllow3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprAllow4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprAllow5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprAllow6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpMH1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpMH2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpMH3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpMH4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpMH5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpMH6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpov1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpov2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpov3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpov4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpov5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpov6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpEq1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpEq2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpEq3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpEq4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpEq5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpEq6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGPRU/A"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpUT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRefR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRHr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRALO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRALB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRALSign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPNoPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPINTPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPINTRT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNoSlv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLSch"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSalesBonus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSalesIndex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSalesIndDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprSalesIndYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpSalesIndRt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpSalesYield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpSalesSpread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpBonus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpIndex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpIndDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGprIndYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpYield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpSpread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFSLTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpBnkNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGEngWarranty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGDTWarranty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PilotRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PilotRevdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.POVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.POSentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO2Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO2SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO3Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO3SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO4Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO4SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO5Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO5SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO6Vendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO6SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.M-T-MExt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.M-T-MExt1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.M-T-MExt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.M-T-MExt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.MaintProv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.MaintTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FAFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SalesFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.CustPO1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.CustPO2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.CustPO3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.CustPO4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpIntBnk"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpCommInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGPRHU/A"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpHUT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpmint"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrphrnt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpTermPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpExtFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpExTxFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRSLTPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpFA4Assign"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpCustPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpSynPenalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpSynRebate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpWHPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SynMemo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.MTMRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.MTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.UnitPctOver"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.MilesOver1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.MilesOver2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.MilesOver3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.MilesOverPct1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.MilesOverPct2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.MilesOverPct3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.HoursOver1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.HoursOver2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.HoursOver3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.HoursOverPct1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.HoursOverPct2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.HoursOverPct3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PONum1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PONum2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PONum3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PONum4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchBnkClNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.BLCDACDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.[CLIENT GROUP]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.GroupAddedDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.UnitsInGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BPChgLO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Strips_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BPChgLO_Up"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCDSetBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BPChgLO_Down"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BPChgSynd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseExtOption1Term"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BPChgSynd_Up"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BrokerFee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ProgressPymtAgmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BPChgSynd_Down"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LRFGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BuyOutTBD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BillToAddressID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BuyOutPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.EarlyTermOptional"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BuyOutNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.CapLeaseBuyOut"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.EarlyTermNote"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ExchangeITGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.OmnibusAgmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ExchangeITMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ExchangeITType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ThirdPartyPayAssgnmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Fee_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FMVBuyOutOptional"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SalePrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FundDate_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.RebateUnitProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.CustomDocumentNameYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.IndexRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePymtForGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LoanType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LoanSecurityAgmtYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LoanDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseTermNotice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.CustomDocumentName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LoanAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.OrigEquipCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PerDiemActual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SaleDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PerDiemProjected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseExtOption1LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ScheduleGroupPaymentTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeaseExtOption1Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.HoldbackResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.RGDRepurchaseTN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.HoldbackConversion"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.RGDTerminationYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.IDFMileageYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FAServicer"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FAInvoicing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FACollecting"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FA3rdPartyPayingAgent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.DateModified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LPAFCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ResidualGuaranteeDocYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SFAExchangeITYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.RGDCasualtyYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.RGDExchangeITYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.ShortFallAgreementYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SFACasualtyYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SFARepurchaseTN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SFATerminationYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.IndemnificationDocYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.IDFLateFeesYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LesseeDefaults"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LesseeObligations"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1009"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1612
    Bottom =872
    Left =-1
    Top =-1
    Right =1580
    Bottom =481
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
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
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
End
