SELECT
  DISTINCT MnRData.UnitID,
  MnRData.VIN,
  MnRData.LogActivityDate AS PrevLogActivityDate,
  MnRData.MilesLTD AS PrevMilesLTD,
  MnRData.HoursLTD AS PrevHoursLTD,
  MnRData.LiftGateCyclesLTD AS PrevLiftGateLTD
FROM
  MnRData
WHERE
  (
    (
      (MnRData.LogActivityDate)= #3/1/2020#
    )
  );
