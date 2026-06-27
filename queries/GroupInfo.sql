SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.MLOrig,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.MLNo,
  vw_SixKeys.FAID,
  SchGrp.SGrpID,
  SchGrp.unitsingroup,
  IIf(
    IsNull([SgrpPlb]),
    No,
    [SgrpPlb]
  ) AS PLB,
  SchGrp.Assigned_Indicator,
  SchGrp.Assignee_Bank,
  SchGrp.AssignmentDate,
  SchGrp.BLCD,
  SchGrp.FundDate,
  SchGrp.LeaseType,
  SchGrp.LeaseTermPrimary,
  SchGrp.BLCDSetBy,
  SchGrp.SGrpNO,
  SchGrp.GroupAddedDate,
  SchGrp.PerDiemStartsOn,
  Lookup_LeaseStartsOn.StartsOn AS PDStart,
  SchGrp.SgrpCommInd,
  SchGrp.OrigEquipCost,
  SchGrp.LeasePymtForGroup,
  SchGrp.ExchangeITGroup,
  SchGrp.SgrpDesc
FROM
  (
    SchGrp
    INNER JOIN vw_SixKeys ON SchGrp.SGrpID = vw_SixKeys.SGrpID
  )
  LEFT JOIN Lookup_LeaseStartsOn ON SchGrp.PerDiemStartsOn = Lookup_LeaseStartsOn.ID
GROUP BY
  vw_SixKeys.clientID,
  vw_SixKeys.MLOrig,
  vw_SixKeys.MLID,
  vw_SixKeys.SchID,
  vw_SixKeys.MLNo,
  vw_SixKeys.FAID,
  SchGrp.SGrpID,
  SchGrp.unitsingroup,
  IIf(
    IsNull([SgrpPlb]),
    No,
    [SgrpPlb]
  ),
  SchGrp.Assigned_Indicator,
  SchGrp.Assignee_Bank,
  SchGrp.AssignmentDate,
  SchGrp.BLCD,
  SchGrp.FundDate,
  SchGrp.LeaseType,
  SchGrp.LeaseTermPrimary,
  SchGrp.BLCDSetBy,
  SchGrp.SGrpNO,
  SchGrp.GroupAddedDate,
  SchGrp.PerDiemStartsOn,
  Lookup_LeaseStartsOn.StartsOn,
  SchGrp.SgrpCommInd,
  SchGrp.OrigEquipCost,
  SchGrp.LeasePymtForGroup,
  SchGrp.ExchangeITGroup,
  SchGrp.SgrpDesc
ORDER BY
  vw_SixKeys.SchID DESC,
  vw_SixKeys.MLNo;
