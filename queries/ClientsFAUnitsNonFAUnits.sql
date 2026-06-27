SELECT
  DISTINCT ClientsFAUnits.ClientGroupName,
  ClientsFAUnits.ClientShortName,
  ClientsFAUnitsFirstOrder.ActualDeliveryDate AS FirstDeliveryDate,
  Sum(ClientsFAUnits.FAUnits) AS SumOfFAUnits,
  Max(
    (
      CInt(
        Nz([NonFAUnits], 0)
      )
    )
  ) AS NotFAUnits,
  Sum([FAUnits])+ Nz(
    Max([NonFAUnits])
  ) AS TotalUnits
FROM
  (
    ClientsFAUnitsFirstOrder
    LEFT JOIN ClientsFAUnits ON ClientsFAUnitsFirstOrder.clientID = ClientsFAUnits.clientID
  )
  LEFT JOIN ClientsNonFAUnits ON ClientsFAUnitsFirstOrder.clientID = ClientsNonFAUnits.clientID
GROUP BY
  ClientsFAUnits.ClientGroupName,
  ClientsFAUnits.ClientShortName,
  ClientsFAUnitsFirstOrder.ActualDeliveryDate
HAVING
  (
    (
      (ClientsFAUnits.ClientGroupName) Is Not Null
    )
  );
