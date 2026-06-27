SELECT
  Units.UnitID,
  Units.GroupREF AS GroupID,
  Units.UnitPDStartDate AS PerdiemStartDate,
  Nz([UNITPDRENT], 0) AS PerdiemRate,
  DateDiff(
    "d",
    [UnitPDStartDate],
    Nz([BLCD], [BLCD_Projected])-1
  ) AS PerdiemDays,
  IIf(
    Nz([PerdiemDays], 0)< 1,
    0,
    [PerdiemDays] * [UnitPDRent]
  ) AS PerdiemTotal,
  IIf(
    IsNull([schgrp].[BLCD]),
    "Projected",
    "Actual"
  ) AS PDType,
  SchGrp.FundDate_Projected,
  SchGrp.BLCD,
  IIf(
    IsNull([schgrp].[FundDate]),
    [schgrp].[FundDate_Projected],
    [schgrp].[FundDate]
  ) AS FundDate,
  SchGrp.BLCD_Projected
FROM
  Units
  INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
WHERE
  (
    (
      (Units.UnitPDStartDate) Is Not Null
    )
  );
