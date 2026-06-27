SELECT
  Clients.ClientShNm,
  MstrLease.MLOrig,
  Schedule.SchExSyn,
  IIf(
    IsNull([MLBACK]),
    IIf(
      IsNull([mlsent]),
      Null,
      "Sent"
    ),
    "Complete"
  ) AS [ML Status],
  Schedule.SchID,
  Schedule.SchNo,
  Schedule.SchIn,
  Schedule.SchOut,
  IIf(
    IsNull([schin]),
    IIf(
      IsNull([schout]),
      Null,
      "Sent"
    ),
    "Complete"
  ) AS [Sch Status],
  IIf(
    IsNull([schin]),
    IIf(
      IsNull([schout]),
      Null,
      "Sent"
    ),
    "Complete"
  ) AS [Sch Status2],
  Schedule.SchFAID,
  Schedule.SchShortDesc,
  MstrLease.MLInsRec,
  MstrLease.MLInsExp,
  MstrLease.MLINSAccpt,
  Vendors_1.VendorSNm,
  SchGrp.FundDate_Projected,
  SchGrp.POSentDt,
  Vendors_2.VendorSNm,
  SchGrp.PO2SentDt,
  Vendors_3.VendorSNm,
  SchGrp.PO3SentDt,
  SchGrp.UnitsInGroup,
  SchGrp.RebateUnitProjected,
  Schedule.SchPDInvDt,
  SchGrp.FundDate,
  SchGrp.LoanBank,
  SchGrp.LoanType,
  SchGrp.SGrpFABill,
  Schedule.PilotRev,
  Schedule.PilotRevdt,
  [max del2].ProjSyn,
  SchGrp.AssignmentDate,
  SchGrp.SGRPNoPD,
  Schedule.SchPDInv,
  PDBillq.MaxOfBillSent AS [PD Billed],
  Schedule.IntLease,
  SchGrp.Assigned_Indicator,
  SchGrp.CustPO1,
  SchGrp.CustPO2,
  SchGrp.CustPO3,
  SchGrp.CustPO4,
  SchGrp.SGrpNO,
  SchGrp.OrigEquipCost,
  SchGrp.Assignee_Bank,
  IIf(
    IsNull([FundDate_Projected]),
    [projsyn],
    [FundDate_Projected]
  ) AS RptProjSyn,
  IIf(
    IsNull([BLCD_Projected]),
    DateAdd("m", 1, [projsyn])-1,
    [BLCD_Projected]
  ) AS rptBlcd,
  [max del2].AccptUnit,
  [max del2].DocsAccpt,
  SchGrp.sgrpnopd,
  SchGrp.BLCD,
  [max del2].MaxOfUnitBuildDate,
  IIf(
    IsNull([sgrpfaid]),
    [schfaid],
    [sgrpfaid]
  ) AS LFaid,
  SchGrp.SGRPPLB,
  SchGrp.FundDate,
  Schedule.SchExSyn
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN (
        Schedule
        LEFT JOIN PDBillq ON Schedule.SchID = PDBillq.BillSchRef
      ) ON MstrLease.MLID = Schedule.SchMLRef
    )
    LEFT JOIN (
      (
        (
          (
            SchGrp
            LEFT JOIN Vendors AS Vendors_1 ON SchGrp.POVendor = Vendors_1.VendorID
          )
          LEFT JOIN Vendors AS Vendors_3 ON SchGrp.PO3Vendor = Vendors_3.VendorID
        )
        LEFT JOIN Vendors AS Vendors_2 ON SchGrp.PO2Vendor = Vendors_2.VendorID
      )
      LEFT JOIN [max del2] ON SchGrp.SGrpID = [max del2].SGrpID
    ) ON Schedule.SchID = SchGrp.SGrpSchID
  )
  LEFT JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (SchGrp.LeaseType)<> "r"
      Or (SchGrp.LeaseType)<> "R"
    )
  )
GROUP BY
  Clients.ClientShNm,
  MstrLease.MLOrig,
  IIf(
    IsNull([MLBACK]),
    IIf(
      IsNull([mlsent]),
      Null,
      "Sent"
    ),
    "Complete"
  ),
  Schedule.SchID,
  Schedule.SchNo,
  Schedule.SchIn,
  Schedule.SchOut,
  IIf(
    IsNull([schin]),
    IIf(
      IsNull([schout]),
      Null,
      "Sent"
    ),
    "Complete"
  ),
  Schedule.SchFAID,
  Schedule.SchShortDesc,
  MstrLease.MLInsRec,
  MstrLease.MLInsExp,
  MstrLease.MLINSAccpt,
  Vendors_1.VendorSNm,
  SchGrp.FundDate_Projected,
  SchGrp.POSentDt,
  Vendors_2.VendorSNm,
  SchGrp.PO2SentDt,
  Vendors_3.VendorSNm,
  SchGrp.PO3SentDt,
  SchGrp.UnitsInGroup,
  SchGrp.RebateUnitProjected,
  Schedule.SchPDInvDt,
  SchGrp.LoanBank,
  SchGrp.LoanType,
  SchGrp.SGrpFABill,
  Schedule.PilotRev,
  Schedule.PilotRevdt,
  [max del2].ProjSyn,
  SchGrp.AssignmentDate,
  SchGrp.SGRPNoPD,
  Schedule.SchPDInv,
  PDBillq.MaxOfBillSent,
  Schedule.IntLease,
  SchGrp.Assigned_Indicator,
  SchGrp.CustPO1,
  SchGrp.CustPO2,
  SchGrp.CustPO3,
  SchGrp.CustPO4,
  SchGrp.SGrpNO,
  SchGrp.OrigEquipCost,
  SchGrp.Assignee_Bank,
  IIf(
    IsNull([FundDate_Projected]),
    [projsyn],
    [FundDate_Projected]
  ),
  IIf(
    IsNull([BLCD_Projected]),
    DateAdd("m", 1, [projsyn])-1,
    [BLCD_Projected]
  ),
  [max del2].AccptUnit,
  [max del2].DocsAccpt,
  SchGrp.sgrpnopd,
  SchGrp.BLCD,
  [max del2].MaxOfUnitBuildDate,
  IIf(
    IsNull([sgrpfaid]),
    [schfaid],
    [sgrpfaid]
  ),
  SchGrp.SGRPPLB,
  SchGrp.FundDate,
  Schedule.SchExSyn,
  SchGrp.FundDate,
  Clients.clientID,
  SchGrp.FundDate,
  IIf(
    IsNull([schin]),
    IIf(
      IsNull([schout]),
      Null,
      "Sent"
    ),
    "Complete"
  ),
  SchGrp.SGrpNO,
  Schedule.SchExSyn,
  Clients.ClientGroupID,
  SchGrp.RebateUnitProjected
HAVING
  (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (SchGrp.LoanBank) Is Null
    )
    AND (
      (Clients.clientID)<> 4302
    )
    AND (
      (SchGrp.FundDate) Is Null
      And (SchGrp.FundDate) Is Null
    )
    AND (
      (Schedule.SchExSyn)= False
    )
    AND (
      (Clients.ClientGroupID)<> 18
    )
  )
  OR (
    (
      (MstrLease.MLOrig)= "fa"
    )
    AND (
      (SchGrp.LoanBank) Is Not Null
    )
    AND (
      (SchGrp.LoanType) Is Not Null
      And (SchGrp.LoanType)<> "Perm"
    )
    AND (
      (Clients.clientID)<> 4302
    )
    AND (
      (SchGrp.FundDate) Is Null
      And (SchGrp.FundDate) Is Null
    )
    AND (
      (Schedule.SchExSyn)= False
    )
    AND (
      (Clients.ClientGroupID)<> 18
    )
  )
ORDER BY
  Clients.ClientShNm,
  Schedule.SchNo;
