SELECT
  ClientGroups.ClientGroupName,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig,
  Count(vw_SixKeys.AssetID) AS FAUnits,
  vw_SixKeys.clientID
FROM
  vw_SixKeys
  INNER JOIN ClientGroups ON vw_SixKeys.ClientGroupId = ClientGroups.ClientGroupId
WHERE
  (
    (
      (vw_SixKeys.Type)= "Tractor"
      Or (vw_SixKeys.Type)= "Straight Truck"
      Or (vw_SixKeys.Type)= "Van"
    )
  )
GROUP BY
  ClientGroups.ClientGroupName,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig,
  vw_SixKeys.clientID,
  vw_SixKeys.Status
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "fa"
    )
    AND (
      (vw_SixKeys.Status)= "A"
      Or (vw_SixKeys.Status)= "TR"
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName;
