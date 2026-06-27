SELECT
  LeaseReport_Syndicated_V_OnOrder.*
FROM
  LeaseReport_Syndicated_V_OnOrder
WHERE
  (
    (
      (
        LeaseReport_Syndicated_V_OnOrder.OrderedYear
      )>= 2014
    )
    And (
      (
        LeaseReport_Syndicated_V_OnOrder.SYDYear
      )= 0
    )
  )
  Or (
    (
      (
        LeaseReport_Syndicated_V_OnOrder.OrderedYear
      )= 0
    )
    And (
      (
        LeaseReport_Syndicated_V_OnOrder.SYDYear
      )>= 2014
    )
  )
ORDER BY
  LeaseReport_Syndicated_V_OnOrder.OrderedYear DESC,
  LeaseReport_Syndicated_V_OnOrder.SYDYear DESC,
  LeaseReport_Syndicated_V_OnOrder.UnitType;
