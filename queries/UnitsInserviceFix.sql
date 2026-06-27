SELECT
  Units.UnitID,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLNo,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  Units.UnitAcceptDate,
  Units.UnitPDStartDate,
  Units.UnitInServ,
  Units.UnitStatus,
  Units.UnitInServInd
FROM
  Units
  INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
WHERE
  (
    (
      (Units.UnitAcceptDate) Is Not Null
    )
    AND (
      (Units.UnitPDStartDate) Is Null
    )
    AND (
      (Units.UnitInServ) Is Null
    )
    AND (
      (Units.UnitStatus)= "A"
    )
  )
  OR (
    (
      (Units.UnitAcceptDate) Is Null
    )
    AND (
      (Units.UnitPDStartDate) Is Not Null
    )
    AND (
      (Units.UnitInServ) Is Null
    )
    AND (
      (Units.UnitStatus)= "A"
    )
  )
  OR (
    (
      (Units.UnitAcceptDate) Is Not Null
    )
    AND (
      (Units.UnitPDStartDate) Is Not Null
    )
    AND (
      (Units.UnitInServ) Is Null
    )
    AND (
      (Units.UnitStatus)= "A"
    )
  )
ORDER BY
  vw_SixKeys.ClientShortName;
