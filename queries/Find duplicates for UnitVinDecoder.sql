SELECT
  UnitVinDecoder.UnitID,
  UnitVinDecoder.DateDecoded,
  UnitVinDecoder.status,
  UnitVinDecoder.ClientID,
  UnitVinDecoder.VIN,
  UnitVinDecoder.make,
  UnitVinDecoder.modelyear,
  UnitVinDecoder.INDX
FROM
  UnitVinDecoder
WHERE
  (
    (
      (UnitVinDecoder.UnitID) In (
        SELECT
          [UnitID]
        FROM
          [UnitVinDecoder] As Tmp
        GROUP BY
          [UnitID]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  UnitVinDecoder.UnitID;
