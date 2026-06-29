SELECT
  Last(Six.Vin) AS VIN,
  Six.Type,
  V.VinGroup,
  [vw_SixKeys].[ClientID] AS ClientID,
  Max(Six.Myear) AS ModelYear
FROM
  (
    vw_SixKeys AS Six
    RIGHT JOIN [SELECT ClientID, VinGroup FROM vw_SixKeys GROUP BY ClientID, VinGroup HAVING (vw_SixKeys].[[ClientID]
  )= 4229] AS V ON Six.VinGRoup = V.VinGroup
)
LEFT JOIN UnitVinDetails ON Six.VIN = UnitVinDetails.VIN
WHERE
  (
    (
      (
        Len([six.VIN])
      )= 17
    )
    AND (
      (UnitVinDetails.status)= "0"
    )
  )
GROUP BY
  Six.Type,
  V.VinGroup,
  [vw_SixKeys].[ClientID];
