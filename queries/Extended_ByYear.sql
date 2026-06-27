SELECT
  UnitExitOption.UnitID,
  UnitExitOption.OptionTypeDesc,
  UnitExitOption.EndOrActualDate,
  UnitExitOption.OrginalEOLDate,
  CInt(
    Nz(
      DatePart("yyyy", [DatePosted]),
      0
    )
  ) AS EXTendYear,
  UnitExitOption.Rent,
  UnitExitOption.TermMonths
FROM
  UnitExitOption
WHERE
  (
    (
      (UnitExitOption.OptionTypeID)= 3
    )
  )
ORDER BY
  UnitExitOption.UnitID;
