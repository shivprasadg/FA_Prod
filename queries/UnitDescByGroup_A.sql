SELECT
  Units.GroupREF,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& " " & [UnitType] AS [Desc],
  Count(Units.UnitID) AS Units,
  vw_SixKeys.MLOrig
FROM
  Units
  INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
WHERE
  (
    (
      (Units.UnitStatus)= "A"
      Or (Units.UnitStatus)= "B"
      Or (Units.UnitStatus)= "C"
      Or (Units.UnitStatus)= "O"
      Or (Units.UnitStatus)= "S"
    )
  )
GROUP BY
  Units.GroupREF,
  vw_SixKeys.MLOrig,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& " " & [UnitType]
HAVING
  (
    (
      (vw_SixKeys.MLOrig)= "fa"
    )
  );
