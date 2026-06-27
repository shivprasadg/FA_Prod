dbMemo "SQL" ="SELECT Clients.ClientShNm, MstrLease.MLOrig, Schedule.SchExSyn, IIf(IsNull([MLBA"
    "CK]),IIf(IsNull([mlsent]),Null,\"Sent\"),\"Complete\") AS [ML Status], Schedule."
    "SchID, Schedule.SchNo, Schedule.SchIn, Schedule.SchOut, IIf(IsNull([schin]),IIf("
    "IsNull([schout]),Null,\"Sent\"),\"Complete\") AS [Sch Status], IIf(IsNull([schin"
    "]),IIf(IsNull([schout]),Null,\"Sent\"),\"Complete\") AS [Sch Status2], Schedule."
    "SchFAID, Schedule.SchShortDesc, MstrLease.MLInsRec, MstrLease.MLInsExp, MstrLeas"
    "e.MLINSAccpt, Vendors_1.VendorSNm, SchGrp.FundDate_Projected, SchGrp.POSentDt, V"
    "endors_2.VendorSNm, SchGrp.PO2SentDt, Vendors_3.VendorSNm, SchGrp.PO3SentDt, Sch"
    "Grp.UnitsInGroup, SchGrp.RebateUnitProjected, Schedule.SchPDInvDt, SchGrp.FundDa"
    "te, SchGrp.LoanBank, SchGrp.LoanType, SchGrp.SGrpFABill, Schedule.PilotRev, Sche"
    "dule.PilotRevdt, [max del2].ProjSyn, SchGrp.AssignmentDate, SchGrp.SGRPNoPD, Sch"
    "edule.SchPDInv, PDBillq.MaxOfBillSent AS [PD Billed], Schedule.IntLease, SchGrp."
    "Assigned_Indicator, SchGrp.CustPO1, SchGrp.CustPO2, SchGrp.CustPO3, SchGrp.CustP"
    "O4, SchGrp.SGrpNO, SchGrp.OrigEquipCost, SchGrp.Assignee_Bank, IIf(IsNull([FundD"
    "ate_Projected]),[projsyn],[FundDate_Projected]) AS RptProjSyn, IIf(IsNull([BLCD_"
    "Projected]),DateAdd(\"m\",1,[projsyn])-1,[BLCD_Projected]) AS rptBlcd, [max del2"
    "].AccptUnit, [max del2].DocsAccpt, SchGrp.sgrpnopd, SchGrp.BLCD, [max del2].MaxO"
    "fUnitBuildDate, IIf(IsNull([sgrpfaid]),[schfaid],[sgrpfaid]) AS LFaid, SchGrp.SG"
    "RPPLB, SchGrp.FundDate, Schedule.SchExSyn\015\012FROM (((Clients INNER JOIN Mstr"
    "Lease ON Clients.clientID = MstrLease.Client) INNER JOIN (Schedule LEFT JOIN PDB"
    "illq ON Schedule.SchID = PDBillq.BillSchRef) ON MstrLease.MLID = Schedule.SchMLR"
    "ef) LEFT JOIN ((((SchGrp LEFT JOIN Vendors AS Vendors_1 ON SchGrp.POVendor = Ven"
    "dors_1.VendorID) LEFT JOIN Vendors AS Vendors_3 ON SchGrp.PO3Vendor = Vendors_3."
    "VendorID) LEFT JOIN Vendors AS Vendors_2 ON SchGrp.PO2Vendor = Vendors_2.VendorI"
    "D) LEFT JOIN [max del2] ON SchGrp.SGrpID = [max del2].SGrpID) ON Schedule.SchID "
    "= SchGrp.SGrpSchID) LEFT JOIN Units ON SchGrp.SGrpID = Units.GroupREF\015\012WHE"
    "RE (((SchGrp.LeaseType)<>\"r\" Or (SchGrp.LeaseType)<>\"R\"))\015\012GROUP BY Cl"
    "ients.ClientShNm, MstrLease.MLOrig, IIf(IsNull([MLBACK]),IIf(IsNull([mlsent]),Nu"
    "ll,\"Sent\"),\"Complete\"), Schedule.SchID, Schedule.SchNo, Schedule.SchIn, Sche"
    "dule.SchOut, IIf(IsNull([schin]),IIf(IsNull([schout]),Null,\"Sent\"),\"Complete\""
    "), Schedule.SchFAID, Schedule.SchShortDesc, MstrLease.MLInsRec, MstrLease.MLInsE"
    "xp, MstrLease.MLINSAccpt, Vendors_1.VendorSNm, SchGrp.FundDate_Projected, SchGrp"
    ".POSentDt, Vendors_2.VendorSNm, SchGrp.PO2SentDt, Vendors_3.VendorSNm, SchGrp.PO"
    "3SentDt, SchGrp.UnitsInGroup, SchGrp.RebateUnitProjected, Schedule.SchPDInvDt, S"
    "chGrp.LoanBank, SchGrp.LoanType, SchGrp.SGrpFABill, Schedule.PilotRev, Schedule."
    "PilotRevdt, [max del2].ProjSyn, SchGrp.AssignmentDate, SchGrp.SGRPNoPD, Schedule"
    ".SchPDInv, PDBillq.MaxOfBillSent, Schedule.IntLease, SchGrp.Assigned_Indicator, "
    "SchGrp.CustPO1, SchGrp.CustPO2, SchGrp.CustPO3, SchGrp.CustPO4, SchGrp.SGrpNO, S"
    "chGrp.OrigEquipCost, SchGrp.Assignee_Bank, IIf(IsNull([FundDate_Projected]),[pro"
    "jsyn],[FundDate_Projected]), IIf(IsNull([BLCD_Projected]),DateAdd(\"m\",1,[projs"
    "yn])-1,[BLCD_Projected]), [max del2].AccptUnit, [max del2].DocsAccpt, SchGrp.sgr"
    "pnopd, SchGrp.BLCD, [max del2].MaxOfUnitBuildDate, IIf(IsNull([sgrpfaid]),[schfa"
    "id],[sgrpfaid]), SchGrp.SGRPPLB, SchGrp.FundDate, Schedule.SchExSyn, SchGrp.Fund"
    "Date, Clients.clientID, SchGrp.FundDate, IIf(IsNull([schin]),IIf(IsNull([schout]"
    "),Null,\"Sent\"),\"Complete\"), SchGrp.SGrpNO, Schedule.SchExSyn, Clients.Client"
    "GroupID, SchGrp.RebateUnitProjected\015\012HAVING (((MstrLease.MLOrig)=\"fa\") A"
    "ND ((SchGrp.LoanBank) Is Null) AND ((Clients.clientID)<>4302) AND ((SchGrp.FundD"
    "ate) Is Null And (SchGrp.FundDate) Is Null) AND ((Schedule.SchExSyn)=False) AND "
    "((Clients.ClientGroupID)<> 18)) OR (((MstrLease.MLOrig)=\"fa\") AND ((SchGrp.Loa"
    "nBank) Is Not Null) AND ((SchGrp.LoanType) Is Not Null And (SchGrp.LoanType)<>\""
    "Perm\") AND ((Clients.clientID)<>4302) AND ((SchGrp.FundDate) Is Null And (SchGr"
    "p.FundDate) Is Null) AND ((Schedule.SchExSyn)=False) AND ((Clients.ClientGroupID"
    ")<> 18))\015\012ORDER BY Clients.ClientShNm, Schedule.SchNo;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbInteger "RowHeight" ="330"
Begin
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbInteger "ColumnWidth" ="1815"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpno"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1155"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Schedule.SchFAID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1170"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNoUnits"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="Schedule.SchShortDesc"
        dbInteger "ColumnWidth" ="4290"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="SchGrp.[SGRPREBATE/UNIT]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="ML Status"
        dbInteger "ColumnWidth" ="1305"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sch Status"
        dbInteger "ColumnWidth" ="1530"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors_1.VendorSNm"
        dbInteger "ColumnWidth" ="2595"
        dbInteger "ColumnOrder" ="13"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors_3.VendorSNm"
        dbInteger "ColumnWidth" ="2595"
        dbInteger "ColumnOrder" ="17"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendors_2.VendorSNm"
        dbInteger "ColumnWidth" ="2595"
        dbInteger "ColumnOrder" ="15"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLInsRec"
        dbInteger "ColumnWidth" ="1290"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLINSAccpt"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PD Billed"
        dbInteger "ColumnWidth" ="1260"
        dbInteger "ColumnOrder" ="20"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchPDInvDt"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpLoanBank"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpLoanTpe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLOrig"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="Schedule.PilotRev"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.PilotRevdt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MstrLease.MLInsExp"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="[max del2].ProjSyn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFunddt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPNoPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchPDInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.IntLease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpFABill"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="30"
    End
    Begin
        dbText "Name" ="SchGrp.POSentDt"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO2SentDt"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.PO3SentDt"
        dbInteger "ColumnWidth" ="1425"
        dbInteger "ColumnOrder" ="18"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpAssInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchIn"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchOut"
        dbInteger "ColumnOrder" ="9"
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
        dbText "Name" ="Sch Status2"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpassignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpPrjFund"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RptProjSyn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="rptBlcd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[max del2].AccptUnit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[max del2].DocsAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpblcd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[max del2].MaxOfUnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LFaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpEquipCost"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchExSyn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1033"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1025"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="POSentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PO2SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PO3SentDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SGrpLoanBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SGrpLoanTpe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SGrpFABill"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjSyn"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SGrpBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SGrpAssInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustPO1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustPO2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustPO3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustPO4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SGrpEquipCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sgrpassignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SGrpPrjFund"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AccptUnit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DocsAccpt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sgrpnopd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sgrpblcd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfUnitBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SGRPPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sgrpfunddt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1002"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assignee_Bank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FundDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.FundDate_Projected"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.UnitsInGroup"
        dbInteger "ColumnOrder" ="25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.RebateUnitProjected"
        dbInteger "ColumnOrder" ="26"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1026"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LoanBank"
        dbInteger "ColumnOrder" ="28"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.AssignmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LoanType"
        dbInteger "ColumnOrder" ="29"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1034"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1022"
        dbLong "AggregateType" ="-1"
    End
End
