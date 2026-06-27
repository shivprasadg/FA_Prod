SELECT
  SchGrp.SGrpSchID AS SchID,
  Schedule.SchNo AS Schedule,
  Format(
    Sum(
      IIf(
        Nz([Units].[DrawDownId], 0)= 0,
        [UnitTotalCost],
        0
      )
    ),
    "$#,###"
  ) AS OpenAmt,
  Format(
    Sum(
      IIf(
        Nz([Units].[DrawDownId], 0)= 0,
        0,
        [unitIFAmt]
      )
    ),
    "$#,###"
  ) AS DrawnAmt,
  Schedule.SchMLRef AS MLID,
  Sum(
    IIf(
      Nz([Units].[DrawDownId], 0)= 0,
      1,
      0
    )
  ) AS OpenUnits,
  Sum(
    IIf(
      Nz([Units].[DrawDownId], 0)= 0,
      0,
      1
    )
  ) AS DrawnUnits,
  Min(Units.UnitIfBank) AS IFBank
FROM
  (
    (
      Units
      INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
    )
    INNER JOIN Schedule ON SchGrp.SGrpSchID = Schedule.SchID
  )
  INNER JOIN UnitTotalOEC ON Units.UnitID = UnitTotalOEC.AssetID
GROUP BY
  SchGrp.SGrpSchID,
  Schedule.SchNo,
  Schedule.SchMLRef;
