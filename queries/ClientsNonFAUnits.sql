SELECT
  ClientGroups.ClientGroupName,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig,
  Count(vw_SixKeys.AssetID) AS NonFAUnits,
  vw_SixKeys.clientID
FROM
  vw_SixKeys
  INNER JOIN ClientGroups ON vw_SixKeys.ClientGroupId = ClientGroups.ClientGroupId
WHERE
  (
    (
      (vw_SixKeys.Status)= "T"
      Or (vw_SixKeys.Status)= "TR"
      Or (vw_SixKeys.Status) Is Null
    )
    AND (
      (vw_SixKeys.Type)= "Tractor"
      Or (vw_SixKeys.Type)= "Straight Truck"
      Or (vw_SixKeys.Type)= "Van"
      Or (vw_SixKeys.Type) Is Null
    )
  )
GROUP BY
  ClientGroups.ClientGroupName,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig,
  vw_SixKeys.clientID
HAVING
  (
    (
      (vw_SixKeys.MLOrig)<> "FA"
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName;
