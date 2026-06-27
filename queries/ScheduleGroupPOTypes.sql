SELECT
  SchGrp.SGrpSchID AS ScheduleID,
  SchGrp.SGrpID AS GroupID,
  IIf(
    [CustPO1] = 0
    And [CustPO2] = 0
    And [CustPO3] = 0
    And [CustPO4] = 0,
    "FA", "Cust"
  ) AS POType,
  IIf(
    IsNull([PONum1]),
    "TBD",
    [PONum1]
  ) AS Ordered,
  IIf(
    [POType] = "Cust",
    "Cust",
    IIf(
      IsNull([POSentDt]),
      "TBD",
      [POSentDt]
    )
  ) AS OrderedDate
FROM
  SchGrp;
