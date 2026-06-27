SELECT
  MstrLease.MLNo,
  Schedule.SchID,
  Schedule.SchNo,
  [forms]![BillingSubForm]![billfrom] AS STDt,
  [forms]![BillingSubForm]![billto] AS enddt,
  Units.unitunitnum,
  Units.unitvin,
  [unityr] & " " & [unitmake] & " " & [unitmodel] & " " & [unittype] AS unitdesc,
  Units.unitvendorpytdate,
  SchGrp.UnitsInGroup,
  IIf(
    [unitvendorpytdate] < [forms]![BillingSubForm]![billfrom],
    [forms]![BillingSubForm]![billfrom],
    [unitvendorpytdate]
  ) AS INtStDt,
  IIf(
    [unitinservind] = "a"
    And [unitinserv] < [forms]![BillingSubForm]![billto],
    [unitinserv] - 1, [forms]![BillingSubForm]![billto]
  ) AS IntEndDt,
  SchGrp.SGRPINTRT,
  Round(
    DateDiff(
      "d",
      [unitvendorpytdate],
      IIf(
        [unitinservind] = "a"
        And [unitinserv] < [forms]![BillingSubForm]![billto],
        [unitinserv], [forms]![BillingSubForm]![billto]
      )
    )* [unitfinalcost] * [SGRPINTRT] / 360,
    2
  ) AS IntAmt,
  Units.UnitInServ,
  Units.UnitInServInd,
  Units.UNITPDRENT,
  IIf(
    IsNull([unitinservind]),
    Null,
    IIf(
      [unitinserv] < [forms]![BillingSubForm]![billfrom],
      [forms]![BillingSubForm]![billfrom],
      [unitinserv]
    )
  ) AS PDStart,
  IIf(
    IsNull([unitinservind])
    Or [unitinservind] <> "a",
    Null,
    [forms]![BillingSubForm]![billto]
  ) AS PDend,
  IIf(
    [unitinservind] <> "a"
    Or IsNull([unitinservind]),
    -1,
    DateDiff(
      "d",
      IIf(
        IsNull([unitinservind]),
        Null,
        IIf(
          [unitinserv] < [forms]![BillingSubForm]![billfrom],
          [forms]![BillingSubForm]![billfrom],
          [unitinserv]
        )
      ),
      IIf(
        IsNull([unitinservind])
        Or [unitinservind] <> "a",
        Null,
        [forms]![BillingSubForm]![billto]
      )
    )
  )+ 1 AS PDDAys,
  (
    IIf(
      [unitinservind] <> "a",
      0,
      DateDiff(
        "d",
        IIf(
          IsNull([unitinservind]),
          Null,
          IIf(
            [unitinserv] < [forms]![BillingSubForm]![billfrom],
            [forms]![BillingSubForm]![billfrom],
            [unitinserv]
          )
        ),
        IIf(
          IsNull([unitinservind])
          Or [unitinservind] <> "a",
          Null,
          [forms]![BillingSubForm]![billto]
        )
      )
    )
  )* [unitpdrent] AS PDRent,
  Units.UnitFinalCost,
  [SGRPINTRT] / 360 AS DailyIntRt,
  Round(
    DateDiff(
      "d",
      [unitvendorpytdate],
      IIf(
        [unitinservind] = "a"
        And [unitinserv] < [forms]![BillingSubForm]![billto],
        [unitinserv], [forms]![BillingSubForm]![billto]
      )
    )
  ) AS IntDays,
  Locations.LocationStreet AS bloc,
  Locations.LAddrLine1 AS bloc2,
  Locations.LAddrLine2 AS BiLo3,
  Locations.LCity AS BCity,
  Locations.LState AS Bst,
  Locations.LZip AS BZip,
  SchGrp.SGrpNO,
  Locations_1.LocationStreet,
  Locations_1.LAddrLine2,
  Locations_1.LAddrLine3,
  Locations_1.LCity,
  Locations_1.LCounty,
  Locations_1.LState,
  Locations_1.LZip,
  Locations_1.LTaxRate,
  Units.UnitTaxExempt,
  Locations_1.LocationName,
  SchGrp.SGRPINTRT,
  SchGrp.SGRPINTPD,
  SchGrp.SGRPNoPD,
  Clients.FAVenorNo,
  Right([unitvin], 6) AS vin4,
  Schedule.SchFSL,
  lookup_States.StTaxType,
  UCase([invfld1desc] & " " & [invfld1]) AS Bill1,
  UCase([invfld2desc] & " " & [invfld2]) AS Bill2,
  Bills.DueonRec,
  IIf(
    [SGrpIntPD] = 0, Null, [intdays] * [DailyIntRt] * [unitfinalcost]
  )+(
    IIf(
      [sgrpnopd] = Yes
      Or [pddays] = 0, 0, [pddays] * [unitpdrent]
    )
  )+(
    IIf(
      [SchFSL] =-1,
      Null,
      (
        IIf(
          [unittaxexempt] = Yes
          Or [sttaxtype] <> "rr",
          0,
          (
            (
              IIf([SGrpIntPD] = 0, Null, 0)
            )+(
              IIf(
                [sgrpnopd] = Yes
                Or [pddays] = 0, 0, [pddays] * [unitpdrent]
              )
            )
          )* [locations_1].[LTaxRate]
        )
      )
    )
  ) AS LineTot
FROM
  (
    (
      (
        (
          Clients
          INNER JOIN Locations ON Clients.clientID = Locations.ClientREF
        )
        INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
      )
      INNER JOIN (
        Schedule
        LEFT JOIN Bills ON Schedule.SchMLRef = Bills.BillSchRef
      ) ON MstrLease.MLID = Schedule.SchMLRef
    )
    INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
  )
  INNER JOIN (
    (
      Units
      INNER JOIN Locations AS Locations_1 ON Units.UnitGarageLocationRef = Locations_1.LocationID
    )
    INNER JOIN lookup_States ON Locations_1.LState = lookup_States.StateAbbrev
  ) ON SchGrp.SGrpID = Units.GroupREF
WHERE
  (
    (
      (Schedule.SchID)= forms!Billing!selsch
    )
    And (
      (Units.unitvendorpytdate) Is Not Null
      And (Units.unitvendorpytdate)<= Forms!BillingSubForm!billto
    )
    And (
      (MstrLease.MLOrig)= "fa"
    )
    And (
      (Units.unitstatus)= "a"
    )
    And (
      (Locations.LocationType)= "billing"
    )
  )
ORDER BY
  Schedule.SchID,
  Schedule.SchNo,
  Units.unitunitnum;
