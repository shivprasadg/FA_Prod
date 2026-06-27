SELECT
  DISTINCT Bills.BillID,
  Bills.BillBank,
  Bills.BillNo AS InvoiceNumber,
  Bills.Billdt,
  Bills.Billfrom,
  Bills.Billto,
  Bills.BillDueDt,
  vw_SixKeys.ClientGroupId,
  vw_SixKeys.Schedule,
  vw_SixKeys.UnitGroup,
  vw_SixKeys.clientID,
  vw_SixKeys.CompanyName,
  vw_SixKeys.MLOrig,
  vw_SixKeys.MLNo,
  vw_SixKeys.AssetID,
  vw_SixKeys.UnitNum,
  Units.UnitStatus,
  vw_SixKeys.VIN,
  [unityr] & " " & [unitmake] & " " & [unitmodel] & IIf(
    Nz([UnitSubType], "")= "",
    "",
    " " & [UnitSubType]
  )& " " & [UnitType] AS [Desc],
  UnitOEC2.UnitCost,
  UnitOEC2.TotalCost,
  Units.UnitTaxExempt AS TaxExempt,
  [Locations].[LocationName] & " - " & [Locations].[LState] AS GaragedLocation,
  lookup_States.StTaxType,
  Locations.LTaxRate,
  Units.UnitAcceptDate AS UnitAccepted,
  SchGrp.BLCD AS BLCD,
  [Locations_1].[LocationStreet] & Chr(13)& Chr(10)& [locations_1].[LAddrLine1] & Chr(13)& Chr(10)& IIf(
    IsNull([locations_1].[LAddrLine2]),
    Null,
    [locations_1].[LAddrLine2] & Chr(13)& Chr(10)
  )& [locations_1].[LCity] & " " & [locations_1].[LState] & " " & [locations_1].[LZip] AS BillToAddress,
  Bank_Assignee.BankAddress,
  "Fleet Advantage, LLC" & Chr(13)& Chr(10)& "401 East Las Olas Blvd. Suite 1720" & Chr(13)& Chr(10)& "Fort Lauderdale, FL 33301" AS FARemit,
  BillingRulesClient.RemitTo,
  IIf(
    [RemitTo] = 2, [FARemit], [BankAddress]
  ) AS RemitAddress,
  CCur(
    IIf(
      [BillingRulesClient].[BillType] > 5,
      [UnitRent], [UnitsRent]
    )
  ) AS Rent,
  IIf(
    [UnitTaxExempt] = True, 0, [locations].[LTaxRate] * [Rent]
  ) AS RentTax,
  IIf(
    [FeeTaxable] = 0
    Or [UnitTaxExempt] = True,
    0,
    (
      [Feevalue] * [locations].[LTaxRate]
    )
  ) AS FeeTax,
  [RentTax] + [FeeTax] AS TaxAmt,
  IIf(
    [ShowContract] = 1, [SchBnkClNO] & "-" & [SgrpBnkNo],
    "hide"
  ) AS BankContract,
  Bills.DueonRec,
  IIf(
    [feeterm] = 0,
    DateAdd(
      "m", vw_SixKeys.[LeaseTermPrimary],
      [BLCD]
    ),
    DateAdd("m", [FeeTerm], [BLCD])
  )-1 AS FeeEndDate,
  CCur(
    IIf(
      [FeeToUse] = 0
      Or [FeeEndDate] < [Billfrom],
      0, [FeeValue]
    )
  ) AS Fee,
  CCur([Rent] + [TaxAmt] + [Fee]) AS LineTotal,
  BillingRules_FeeLabels.RuleString AS FeeLabel,
  1 AS UnitCount
FROM
  (
    (
      (
        (
          (
            (
              (
                Bills
                INNER JOIN (
                  (
                    (
                      vw_SixKeys
                      LEFT JOIN BillingRulesClient ON (
                        vw_SixKeys.SchID = BillingRulesClient.ScheduleId
                      )
                      AND (
                        vw_SixKeys.clientID = BillingRulesClient.ClientId
                      )
                    )
                    INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID
                  )
                  INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
                ) ON Bills.BillSchRef = vw_SixKeys.SchID
              )
              INNER JOIN UnitOEC2 ON vw_SixKeys.AssetID = UnitOEC2.AssetId
            )
            INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
          )
          INNER JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev
        )
        INNER JOIN Schedule ON vw_SixKeys.SchID = Schedule.SchID
      )
      INNER JOIN Locations AS Locations_1 ON Schedule.SchBillAdd = Locations_1.LocationID
    )
    INNER JOIN Bank_Assignee ON SchGrp.Assignee_Bank = Bank_Assignee.ShortName
  )
  LEFT JOIN BillingRules_FeeLabels ON BillingRulesClient.FeeLabel = BillingRules_FeeLabels.RuleValue
WHERE
  (
    (
      (Bills.BillID)= 1476
    )
    AND (
      (vw_SixKeys.clientID)= 4971
    )
    AND (
      (vw_SixKeys.MLOrig)= "FA"
    )
    AND (
      (Units.UnitStatus)= "A"
    )
  );
