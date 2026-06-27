SELECT
  Units.UnitInServ AS InService,
  Units.unitstatus AS Status,
  Units.Axle,
  Units.UNITSRENT AS Rent,
  SchGrp.BLCD AS BLCD,
  Round([LeaseTermPrimary]) AS Term,
  CDate(
    IIf(
      IsNull([BLCD]),
      "1/1/2050",
      (
        DateAdd("m", [LeaseTermPrimary], [BLCD])-1
      )
    )
  ) AS LeaseEndDate,
  SchGrp.ExchangeITMonth AS ExchangetITTerm,
  CDate(
    IIf(
      IsNull([BLCD]),
      "1/1/2050",
      (
        DateAdd("m", [ExchangeITMonth], [BLCD])-1
      )
    )
  ) AS ExchangeITDate,
  CDate(
    IIf(
      IsNull([BLCD]),
      "1/1/2050",
      DateAdd("m", [LeaseTermPrimary], [BLCD])-90
    )
  ) AS NoticeDate,
  IIf(
    DateDiff(
      "m",
      Date(),
      CDate(
        IIf(
          IsNull([BLCD]),
          "1/1/2050",
          (
            DateAdd("m", [ExchangeITMonth], [BLCD])-1
          )
        )
      )
    )<-3,
    0,
    [ExchangeITGroup]
  ) AS ExchangeIT,
  SchGrp.LeaseExtOption1Term AS ExtOption,
  Units.UnitSynResid AS ExpRV,
  SchGrp.Assignee_Bank AS Assignee,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] AS [DESC],
  DateDiff(
    "m",
    Date(),
    CDate(
      IIf(
        IsNull([BLCD]),
        "1/1/2050",
        DateAdd("m", [ExchangeITMonth], [BLCD])
      )
    )
  ) AS MonthsUnitEXIT,
  DateDiff(
    "m",
    Date(),
    CDate(
      IIf(
        IsNull([BLCD]),
        "1/1/2050",
        DateAdd("m", [LeaseTermPrimary], [BLCD])
      )
    )
  ) AS MonthsUnitEOL,
  SchGrp.LeaseExtOption1Term AS ExchangeItMonths
FROM
  (
    SchGrp
    INNER JOIN Schedule ON SchGrp.SGrpSchID = Schedule.SchID
  )
  INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Units.unitstatus)<> "S"
      And (Units.unitstatus)<> "KS"
    )
    AND (
      (
        DateDiff(
          "m",
          Date(),
          CDate(
            IIf(
              IsNull([BLCD]),
              "1/1/2050",
              DateAdd("m", [ExchangeITMonth], [BLCD])
            )
          )
        )
      ) Between -12
      And 25
    )
  );
