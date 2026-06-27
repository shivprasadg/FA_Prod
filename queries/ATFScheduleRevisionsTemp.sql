SELECT
  4 AS RevisionNum,
  #3/30/2026# AS RevisionDate,
  'ljohnson' AS SysUser,
  'Schedule Inserted into ATLAAS' AS RevisionNote,
  Clients.clientID AS LeaseeID,
  Clients.ClientShNm,
  Clients.clientTaxID,
  MstrLease.MLID,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  MstrLease.MLDt,
  Schedule.SchID,
  SchGrp.SGrpID,
  SchGrp.SGrpNO,
  Schedule.ATFID,
  Schedule.ATFSchSpecID,
  SchGrp.ATFGroupSpecID,
  Units.ATFUnitSpecID,
  Count(Units.UnitID) AS Units,
  Max(Units.UnitAcceptDate) AS UnitAcceptDate,
  Max(Units.UnitCofACopy) AS UnitCofACopy,
  Max(Units.UnitCOABack) AS UnitCOABack,
  Units.UnitRegLocationRef,
  Units.UnitTitleLocationRef,
  Max(Units.UnitVendorPytDate) AS UnitVendorPytDate,
  Max(Units.UnitActualDelDate) AS UnitActualDelDate,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeaseTermNotice,
  SchGrp.LeaseType,
  SchGrp.LeaseExtOption1Term,
  SchGrp.ExchangeITType,
  SchGrp.LeasePymtFreq,
  SchGrp.LeasePaidAdvArr,
  SchGrp.SGRPPLB,
  SchGrp.SGrpSalesIndex,
  SchGrp.SGrpSalesIndDt,
  SchGrp.SGprSalesIndYr,
  SchGrp.SgrpSalesIndRt,
  SchGrp.SgrpSalesYield,
  SchGrp.SGrpSalesBonus,
  SchGrp.SGrpIndex,
  SchGrp.SGrpIndDt,
  SchGrp.SGprIndYr,
  SchGrp.IndexRate,
  SchGrp.SgrpYield,
  SchGrp.SGrpBonus,
  SchGrp.SgrpCommInd,
  SchGrp.PerDiemStartsOn,
  SchGrp.BLCD,
  SchGrp.Assignee_Bank,
  Round(
    Avg(Units.UNITRENT),
    2
  ) AS UNITRENT,
  Round(
    Avg(Units.UNITSRENT),
    2
  ) AS UNITSRENT,
  Round(
    Avg(Units.FirstExtRent),
    2
  ) AS FirstExtRent,
  Round(
    Avg(Units.MTMRent),
    2
  ) AS MTMRent,
  Round(
    Avg(Units.UnitSalesResid),
    2
  ) AS UnitSalesResid,
  Round(
    Avg(Units.UnitSynResid),
    2
  ) AS UnitSynResid,
  Round(
    Avg(Units.UnitFinalCost),
    2
  ) AS UnitFinalCost,
  Schedule.FSLType,
  Schedule.ApprovalRequest,
  Schedule.ApprovalReply,
  Schedule.SchNo,
  Schedule.SchShortDesc,
  Schedule.BankId,
  Schedule.SchRALO,
  Schedule.SchRALB,
  Schedule.SchRALSigner,
  Schedule.SchBillAdd,
  Units.VendorREF,
  Max(Units.UnitEstDelDate) AS UnitEstDelDate,
  Round(
    Avg(Units.UNITPDRENT),
    2
  ) AS UNITPDRENT,
  Schedule.SchFAID,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType,
  Units.UnitSubType,
  Units.Axle,
  Round(
    Avg(Units.VendorCost),
    2
  ) AS VendorCost
FROM
  (
    (
      (
        Clients
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  Schedule.SchID = 3028
GROUP BY
  Clients.clientID,
  MstrLease.MLID,
  MstrLease.MLNo,
  MstrLease.MLOrig,
  MstrLease.MLDt,
  Clients.ClientShNm,
  Clients.clientTaxID,
  Schedule.SchID,
  SchGrp.SGrpID,
  SchGrp.SGrpNO,
  Schedule.ATFID,
  Schedule.ATFSchSpecID,
  SchGrp.ATFGroupSpecID,
  Units.ATFUnitSpecID,
  Units.UnitRegLocationRef,
  Units.UnitTitleLocationRef,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeaseTermNotice,
  SchGrp.LeaseType,
  SchGrp.LeaseExtOption1Term,
  SchGrp.ExchangeITType,
  SchGrp.LeasePymtFreq,
  SchGrp.LeasePaidAdvArr,
  SchGrp.SGRPPLB,
  SchGrp.SGrpSalesIndex,
  SchGrp.SGrpSalesIndDt,
  SchGrp.SGprSalesIndYr,
  SchGrp.SgrpSalesIndRt,
  SchGrp.SgrpSalesYield,
  SchGrp.SGrpSalesBonus,
  SchGrp.SGrpIndex,
  SchGrp.SGrpIndDt,
  SchGrp.SGprIndYr,
  SchGrp.IndexRate,
  SchGrp.SgrpYield,
  SchGrp.SGrpBonus,
  SchGrp.SgrpCommInd,
  SchGrp.PerDiemStartsOn,
  SchGrp.BLCD,
  SchGrp.Assignee_Bank,
  Schedule.FSLType,
  Schedule.ApprovalRequest,
  Schedule.ApprovalReply,
  Schedule.SchNo,
  Schedule.SchShortDesc,
  Schedule.BankId,
  Schedule.SchRALO,
  Schedule.SchRALB,
  Schedule.SchRALSigner,
  Schedule.SchBillAdd,
  Units.VendorREF,
  Schedule.SchFAID,
  Units.UnitYr,
  Units.UnitMake,
  Units.UnitModel,
  Units.UnitType,
  Units.UnitSubType,
  Units.Axle
HAVING
  (MstrLease.MLOrig = 'FA');
