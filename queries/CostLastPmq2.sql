SELECT
  Costcomaxpmdt.LocationName,
  Costcomaxpmdt.UnitUnitNum,
  Costcomaxpmdt.UnitVIN,
  Costcomaxpmdt.MaxOfPMDate,
  DateDiff(
    "d",
    [maxofpmdate],
    Date()
  ) AS days,
  Units.unityr,
  Units.unitmake,
  Units.unitmodel,
  Units.unittype
FROM
  Units
  INNER JOIN Costcomaxpmdt ON Units.UnitVIN = Costcomaxpmdt.UnitVIN
WHERE
  (
    (
      (
        DateDiff(
          "d",
          [maxofpmdate],
          Date()
        )
      )> 100
    )
  )
ORDER BY
  Costcomaxpmdt.LocationName,
  Costcomaxpmdt.UnitUnitNum;
