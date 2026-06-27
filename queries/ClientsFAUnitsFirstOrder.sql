SELECT
  ClientGroups.ClientGroupName,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig,
  Count(vw_SixKeys.AssetID) AS Units,
  Min(
    IIf(
      Nz(
        [vw_SixKeys.ActualDeliveryDate],
        "1/1/2999"
      )> [BLCD],
      [BLCD],
      [vw_SixKeys.ActualDeliveryDate]
    )
  ) AS ActualDeliveryDate,
  vw_SixKeys.clientID
FROM
  vw_SixKeys
  INNER JOIN ClientGroups ON vw_SixKeys.ClientGroupId = ClientGroups.ClientGroupId
GROUP BY
  ClientGroups.ClientGroupName,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig,
  vw_SixKeys.clientID
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName;
