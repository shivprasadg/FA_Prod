SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.MLOrig,
  Count(Units.UnitID) AS Units
FROM
  (
    vw_SixKeys
    INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
  )
  INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID
WHERE
  (
    (
      (Units.unitstatus)= "A"
      Or (Units.unitstatus)= "C"
    )
    AND (
      (vw_SixKeys.MLID) Is Not Null
    )
  )
GROUP BY
  vw_SixKeys.clientID,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLOrig
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName;
