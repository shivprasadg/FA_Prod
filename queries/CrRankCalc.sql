SELECT
  Clients.clientID,
  Clients.ClientShNm,
  Clients.clientType,
  IIf(
    [fleetdaycabs] + [fleetsleepers] >= 400,
    0.25,
    IIf(
      [fleetdaycabs] + [fleetsleepers] >= 200,
      0.15, 0
    )
  ) AS Fleet,
  IIf(
    [FLEETDAYCABS] = 0,
    0,
    IIf(
      [FLEETDAYCABSMPY] = 0
      Or [fleetdaycabsmpy] >= 80000,
      0,
      IIf(
        [FLEETDAYCABSMPY] > 70000,
        -0.1,
        IIf(
          [FLEETDAYCABSMPY] > 60000,-0.2,-0.3
        )
      )
    )
  ) AS DCMI,
  IIf(
    [FLEETSLEEPERS] = 0,
    0,
    IIf(
      [FLEETSLEEPERSMPY] = 0
      Or [FLEETSLEEPERSMPY] >= 10000,
      0,-0.3
    )
  ) AS SlpMI,
  IIf(
    IsNull([FleetReplacementCycle]),
    0,
    IIf(
      [FleetReplacementCycle] > 4.99, 0.15,
      0
    )
  ) AS Cycle,
  IIf(
    [clientobc] = "none",
    0.05,
    IIf([clientobc] = "peoplenet", 0.1, 0)
  ) AS OldOBC,
  IIf(
    [creditmoodyrating] = "aaa",
    0.4,
    IIf(
      Left([creditmoodyrating], 2)= "aa",
      0.35,
      IIf(
        Left([creditmoodyrating], 1)= "a",
        0.3,
        IIf(
          Left([creditmoodyrating], 3)= "baa",
          0.25,
          IIf(
            Left([creditmoodyrating], 2)= "BAA",
            0.25,
            IIf(
              Left([creditmoodyrating], 3)= "BBB",
              0.25,
              IIf(
                Left([creditmoodyrating], 2)= "BA",
                0.15,
                IIf(
                  Left([creditmoodyrating], 2)= "BB",
                  0.15,
                  0
                )
              )
            )
          )
        )
      )
    )
  ) AS Moody,
  IIf(
    [CREDITRANK] = "1",
    0.4,
    IIf(
      [CREDITRANK] = "2",
      0.35,
      IIf(
        [CREDITRANK] = "3",
        0.3,
        IIf(
          [CREDITRANK] = "4",
          0.25,
          IIf([CREDITRANK] = "5", 0.15, 0)
        )
      )
    )
  ) AS CreditBank,
  IIf(
    [creditfinance] = "purchase",
    0.15,
    IIf(
      [creditfinance] = "lease",
      0.1,
      IIf([creditfinance] = "FULL",-0.1, 0)
    )
  ) AS Acquisition,
  IIf(
    [creditmaint] = "in-house",
    0.15,
    IIf(
      [creditmaint] = "outsource"
      And [CREDITFINANCE] <> "FULL",
      0.05, 0
    )
  ) AS Maint
FROM
  Clients
WHERE
  (
    (
      (Clients.clientType)<> "vendor"
      And (Clients.clientType)<> "bank"
      And (Clients.clientType)<> "fa"
    )
  )
ORDER BY
  Clients.ClientShNm;
