SELECT
  vw_SixKeys.ClientShortName AS Client,
  Bank_Assignee.ShortName AS Assignee,
  Bank_Assignee.BankName,
  vw_SixKeys.Schedule,
  SchGrp.SGrpNO AS [Group],
  Units.unitunitnum AS UnitNum,
  Units.UnitVIN,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] & " " & [Axle] AS [Desc],
  Locations.LocationName,
  [SchBnkClNO] & IIf(
    Left([SgrpBnkNo], 1)= "-",
    [SgrpBnkNo],
    "-" & [SgrpBnkNo]
  ) AS [Bank Client No],
  IIf(
    [UnitTaxExempt] =-1, "Exempt", "Taxed"
  ) AS TaxExempt,
  lookup_States.StTaxType AS [Tax Type],
  Locations.LTaxRate AS [Tax Rate],
  CCur([unitsrent] + 0) AS Rent,
  CCur(
    Nz([SGrpFSLPmt], 0)
  ) AS FSLFee,
  CCur(
    IIf(
      [unittaxexempt] =-1,
      0,
      Round(
        [locations].[LTaxRate] * [unitsrent],
        2
      )
    )
  ) AS [Tax Amt],
  [FslFee] + [Rent] + [Tax Amt] AS [Line Total]
FROM
  (
    (
      (
        (
          (
            (
              (
                vw_SixKeys
                INNER JOIN Units ON vw_SixKeys.AssetID = Units.UnitID
              )
              INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
            )
            INNER JOIN Locations ON Units.UnitGarageLocationRef = Locations.LocationID
          )
          INNER JOIN Schedule ON vw_SixKeys.SchID = Schedule.SchID
        )
        INNER JOIN Locations AS Locations_1 ON Schedule.SchBillAdd = Locations_1.LocationID
      )
      INNER JOIN lookup_States ON Locations.LState = lookup_States.StateAbbrev
    )
    LEFT JOIN Bank_Assignee ON SchGrp.Assignee_Bank = Bank_Assignee.ShortName
  )
  LEFT JOIN FSLFee_EndDate ON vw_SixKeys.SGrpID = FSLFee_EndDate.GroupId
WHERE
  (
    (
      (vw_SixKeys.SchID)= [Forms]![Schedule_Form]![SchID]
    )
  );
