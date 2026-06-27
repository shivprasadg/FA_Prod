SELECT
  NewInsp.RequestID,
  NewInsp.STATUS,
  NewInsp.ClaimNum,
  NewInsp.iUnitNum,
  NewInsp.iDate,
  NewInsp.INSP,
  NewInsp.LinkAddr,
  NewInsp.iCost,
  NewInsp.VIN,
  NewInsp.TYPE,
  NewInsp.Year,
  NewInsp.MAKE,
  NewInsp.MODEL,
  NewInsp.CITY,
  NewInsp.ST,
  NewInsp.LOCATION
FROM
  (
    SELECT
      CLng(
        Nz([REQUEST ID], 0)
      ) AS RequestID,
      NewDekraInspections.STATUS,
      NewDekraInspections.[CLAIM#] AS ClaimNum,
      NewDekraInspections.[UNIT#] AS iUnitNum,
      CDate(
        Nz(
          [NewDekraInspections.REQUEST], [NewDekraInspections.REQUEST]
        )
      ) AS iDate,
      CLng(
        Nz([INSP #], 0)
      ) AS INSP,
      NewDekraInspections.LinkAddr,
      CCur(
        Nz([PRICE TO CUSTOMER], 0)
      ) AS iCost,
      NewDekraInspections.VIN,
      NewDekraInspections.TYPE,
      NewDekraInspections.Year,
      NewDekraInspections.MAKE,
      NewDekraInspections.MODEL,
      NewDekraInspections.CITY,
      NewDekraInspections.ST,
      NewDekraInspections.LOCATION
    FROM
      NewDekraInspections
    WHERE
      (
        (
          (NewDekraInspections.STATUS)= "Completed"
        )
      )
  ) AS NewInsp
  LEFT JOIN Units ON (
    NewInsp.iUnitNum = Units.UnitUnitNum
  )
  AND (NewInsp.VIN = Units.UnitVIN)
WHERE
  (
    (
      (Units.UnitID) Is Null
    )
  );
