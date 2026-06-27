SELECT
  vw_SixKeys.ClientShortName,
  Count(vw_SixKeys.AssetID) AS UnitTitled,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId,
  Min(Units.UnitTASent) AS FirstTitled,
  Max(Units.UnitTASent) AS LastTitled
FROM
  vw_SixKeys
  INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
    AND (
      (Units.UnitTASent)>= '1/1/' & [AskYear]
      And (Units.UnitTASent)< '12/31/' & [AskYear]
    )
  )
GROUP BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.clientID,
  vw_SixKeys.ClientGroupId;
