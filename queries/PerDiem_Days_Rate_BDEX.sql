SELECT
  S.BDEX,
  S.ClientShortName AS Account,
  PD.AssetID,
  PD.BLCD,
  PD.PerdiemStartDate AS PerDiemStart,
  PD.PerDiemDays,
  PD.RentPD AS PerDiemRent,
  PD.PerDiemAmt,
  PD.PerDiemProjected AS PerDiemProj
FROM
  vw_SixKeys AS S
  INNER JOIN PerDiem_Days_Rate_Unit AS PD ON S.UnitID = PD.AssetID
ORDER BY
  S.BDEX,
  S.ClientShortName;
