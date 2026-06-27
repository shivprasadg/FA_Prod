SELECT
  vw_SixKeys.clientID,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLNo,
  Count(Units.UnitID) AS Units,
  Schedule.SchMLRef AS MLID
FROM
  (
    (
      SchGrp
      INNER JOIN Schedule ON SchGrp.SGrpSchID = Schedule.SchID
    )
    INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
  )
  INNER JOIN vw_SixKeys ON Schedule.SchMLRef = vw_SixKeys.MLID
WHERE
  (
    (
      (Units.unitstatus)= "A"
      Or (Units.unitstatus)= "C"
    )
    AND (
      (
        DateDiff(
          "m",
          Date(),
          CDate(
            IIf(
              IsNull([BLCD]),
              "1/1/2050",
              DateAdd("m", [ExchangeITMonth], [BLCD])
            )
          )
        )
      ) Between -12
      And 25
    )
    AND (
      (vw_SixKeys.MLOrig)= "fa"
    )
  )
GROUP BY
  vw_SixKeys.clientID,
  vw_SixKeys.ClientShortName,
  vw_SixKeys.MLNo,
  Schedule.SchMLRef
ORDER BY
  vw_SixKeys.ClientShortName;
