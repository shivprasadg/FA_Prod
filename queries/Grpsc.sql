SELECT
  Clients.ClientGroupID,
  Clients.ClientShNm,
  IIf(
    [fleetDaycabs] + [fleetSleepers] < 200,
    -0.1,
    0.15 + 0.05 *(
      - Int(
        -(
          (
            [fleetdaycabs] + [fleetsleepers] - 200
          )/ 500
        )
      )
    )
  ) AS Flt,
  IIf(
    [FLEETDAYCABSMPY] = 0,
    0,
    IIf(
      [FLEETDAYCABSMPY] < 60000,
      -0.4,
      IIf(
        [FLEETDAYCABSMPY] < 70000,
        -0.2,
        IIf(
          [FLEETDAYCABSMPY] < 80000,-0.05, 0
        )
      )
    )
  ) AS DCMI,
  IIf(
    [fleetSleepersMPY] = 0,
    0,
    IIf(
      [fleetSleepersMPY] < 70000,
      -0.4,
      IIf(
        [fleetSleepersMPY] < 80000,
        -0.2,
        IIf(
          [fleetSleepersMPY] < 90000,-0.1, 0
        )
      )
    )
  ) AS SlpMI,
  IIf(
    [creditrank] <> "tb"
    And [creditMoodyRating] = "tbd",
    0,
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
                    -0.35
                  )
                )
              )
            )
          )
        )
      )
    )
  ) AS AGRT,
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
          IIf([CREDITRANK] = "5", 0.15,-0.35)
        )
      )
    )
  ) AS FART,
  IIf(
    [creditfinance] = "purchase",
    0.15,
    IIf(
      [creditfinance] = "lease",
      0.1,
      IIf(
        [creditfinance] = "FULL service",
        -0.2, 0
      )
    )
  ) AS AQ,
  IIf(
    [creditfinance] = "full service",
    0,
    IIf(
      [creditmaint] = "in-house",
      0.15,
      IIf(
        [creditmaint] = "outsource"
        And [CREDITFINANCE] <> "FULL Service",
        0.05, 0
      )
    )
  ) AS Maint,
  IIf(
    IsNull([FleetReplacementCycle]),
    0,
    IIf(
      [FleetReplacementCycle] > 3.99, 0.15,
      0
    )
  ) AS Cyc,
  IIf(
    [clientobc] = "None"
    Or [clientobc] = "tbd",
    0, 0.1
  ) AS OBC
FROM
  Clients
WHERE
  (
    (
      (Clients.ClientGroupID) Is Not Null
    )
  )
ORDER BY
  Clients.ClientGroupID,
  Clients.ClientShNm;
