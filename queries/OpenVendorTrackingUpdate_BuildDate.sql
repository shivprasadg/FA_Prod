SELECT
  Units.UnitID,
  OpenVendorDataTemp.UnitBuildDate AS BuildDate,
  Nz(
    [Units].[UnitBuildDate], '1/1/1900'
  ) AS PrevBuildDate
FROM
  OpenVendorDataTemp
  INNER JOIN Units ON OpenVendorDataTemp.AssetID = Units.UnitID
WHERE
  (
    (
      (
        Nz(
          [Units].[UnitBuildDate], '1/1/1900'
        )
      )<> [OpenVendorDataTemp].[UnitBuildDate]
    )
  );
