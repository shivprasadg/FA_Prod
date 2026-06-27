SELECT
  Units.UnitTitleLocationRef,
  Units.UnitGarageLocationRef,
  Units.UnitRegLocationRef,
  Units.UnitDeliveryLocationRef,
  CDbl(
    Nz([units].[CofASigner], 0)
  ) AS CofASigner,
  MstrLease.MLID,
  Schedule.SchID,
  Schedule.SchedDt,
  Units.UnitID,
  Units.unitvin,
  [sgrpno] & [unitdeliverylocationref] & [unitgaragelocationref] & [unittitlelocationref] & [unitregLocationref] & [cofasigner] AS GrpCD,
  Clients.clientID,
  Clients.clientCompanyName,
  Clients.clientCoNm2,
  Clients.clientDBA,
  SchGrp.SGrpID,
  Clients.clientAddress1,
  Clients.clientAddress2,
  Clients.clientCity,
  Clients.clientState,
  Clients.clientZip,
  MstrLease.MLNo,
  MstrLease.MLDt,
  DatePart("yyyy", [mldt]) AS MLYr,
  MonthName(
    DatePart("m", [mldt])
  ) AS MLMo,
  DatePart("d", [mldt]) AS mlday,
  Schedule.SchNo,
  lookup_States.StateLong AS CorpState,
  SchGrp.SGrpNO,
  Right([unitvin], 6) AS vin4,
  UCase([unityr]) AS ModelYear,
  Units.UnitMake,
  [unitmake] & "   " & [unitmodel] AS MakeModel,
  Units.UnitType,
  Units.UnitType AS Type,
  Units.UnitSubType,
  Units.Axle,
  IIf([AXLE] = "TBD", Null, [AXLE] & " ")& [UnitSubtype] & " " & [UnitType] AS [Desc],
  Units.UnitUnitNum,
  Units.UnitAcceptDate,
  Units.Unitstatus,
  SchGrp.SGrpSchID,
  SchGrp.SGRPPLB,
  Units.UNITRENT,
  SchGrp.BLCD_Projected,
  Schedule.SchFSL,
  SchGrp.SGrpFSLPmt,
  Units.UnitModel,
  Units.UNITTAX,
  IIf(
    [units].[UNITRENT] > 0, [units].[UNITRENT],
    [units].[unitSRent]
  ) AS RENT,
  IIf(
    [unitfinalcost] = 0, Null, [unitfinalcost]
  ) AS Unitcost,
  Units.UnitTaxExempt,
  Schedule.SubLeasee,
  Units.UnitSubType,
  UnitLastMileageRead.Mileage,
  UnitLastMileageRead.Hours,
  UnitLastMileageRead.ReadDate,
  Nz(
    [Units].[LegalDescription], [lookup_UnitDesc].[DescriptionLong]
  ) AS DescriptionLong,
  SchGrp.SGRPPLB AS PLB,
  IIf(
    [SGRPPLB] = True,
    Nz([AcqMileage], 0),
    0
  ) AS PLBMileage,
  IIf(
    [SGRPPLB] = True,
    Nz([AcqHours], 0),
    0
  ) AS PLBHours,
  Switch(
    [SGRPPLB] = False,
    Null,
    Nz([PLBMileage], 0)> 0,
    "Miles as of " & Format(
      [Forms]![BulkChange_Form]![PLBAsOfDate],
      "mm/dd/yy"
    ),
    Nz([PLBHours], 0)> 0,
    "Hours",
    Nz([PLBMileage], 0)= 0
    And Nz([PLBHours], 0)= 0,
    "Miles / Hours"
  ) AS PLBLabel,
  SchGrp.SGRPPLB,
  Switch(
    [SGRPPLB] = False,
    Null,
    Nz([PLBMileage], 0)> 0,
    Format([PLBMileage], "#,##0"),
    Nz([PLBHours], 0),
    Format([PLBHours], "#,##0"),
    Nz([PLBMileage], 0)> 0
    And Nz([PLBHours], 0)> 0,
    Format([PLBMileage], "#,##0")& " / " & Format([PLBHours], "#,##0")
  ) AS PLBMilesHours,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary
FROM
  (
    (
      (
        (
          (
            (
              Clients
              INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
            )
            INNER JOIN Schedule ON MstrLease.MLID = Schedule.SchMLRef
          )
          INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
        )
        INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
      )
      LEFT JOIN lookup_States ON Clients.LeaseSTofCorp = lookup_States.StateAbbrev
    )
    LEFT JOIN UnitLastMileageRead ON Units.UnitID = UnitLastMileageRead.UnitID
  )
  INNER JOIN LookUp_UnitDesc ON Units.UnitID = LookUp_UnitDesc.AssetID
WHERE
  (
    (
      (Units.Unitstatus)= "A"
      Or (Units.Unitstatus)= "B"
    )
  )
ORDER BY
  Clients.clientCompanyName,
  Right([unitvin], 6),
  Units.UnitUnitNum;
