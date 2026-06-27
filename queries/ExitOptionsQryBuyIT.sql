SELECT
  DISTINCTROW UnitExitOption.*,
  Units.UnitUnitNum,
  Units_1.UnitPDStartDate,
  DLookUp(
    "ParkDateProjected", "RMKT", "UnitRef=" & [Units].[UnitID]
  ) AS ParkDateProjected,
  CCur(
    Nz(
      DLookUp(
        "AdjBookValue", "RMKT", "UnitRef=" & [Units].[UnitID]
      ),
      0
    )
  ) AS AdjBookValue,
  CCur(
    Nz(
      DLookUp(
        "TargetSale", "RMKT", "UnitRef=" & [Units].[UnitID]
      ),
      0
    )
  ) AS TargetSale,
  CCur(
    Nz(
      DLookUp(
        "BuyOutCost", "RMKT", "UnitRef=" & [Units].[UnitID]
      ),
      0
    )
  ) AS BuyOutCost,
  CCur(
    Nz(
      DLookUp(
        "TermPenalty", "RMKT", "UnitRef=" & [Units].[UnitID]
      ),
      0
    )
  ) AS TermPenalty
FROM
  (
    (
      EXITOptions_WorkTable
      INNER JOIN UnitExitOption ON EXITOptions_WorkTable.UnitID = UnitExitOption.UnitId
    )
    INNER JOIN Units ON UnitExitOption.UnitID = Units.UnitID
  )
  LEFT JOIN Units AS Units_1 ON UnitExitOption.ReplacedByUnitID = Units_1.UnitID
WHERE
  (
    (
      (UnitExitOption.OptionTypeID)= 0
    )
  );
