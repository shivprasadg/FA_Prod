SELECT
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.MLOrig,
  Count(vw_SixKeys.AssetID) AS ActualUnits,
  SchGrp.UnitsInGroup,
  IIf(
    [UnitsInGroup] <> Count([AssetID]),
    "Y",
    Null
  ) AS Problem
FROM
  vw_SixKeys
  INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
WHERE
  (
    (
      (vw_SixKeys.SubType)<> "Auto Hauler"
    )
  )
GROUP BY
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.MLOrig,
  SchGrp.UnitsInGroup
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
    AND (
      (
        IIf(
          [UnitsInGroup] <> Count([AssetID]),
          "Y",
          Null
        )
      )= "Y"
    )
  );
