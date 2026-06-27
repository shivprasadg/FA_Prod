INSERT INTO UnitExitOption (
  UnitID, UnitGroupID, ProfileBLCD,
  ProfileTerm, ProfileEOLDate, ProfileType,
  RentLO, RentPerdiem, RentSynd, RentActive,
  ResidualSynd, ResidualLO, AssigneeBankId,
  RentMTM, NoticeReqMonths, UnitOEC,
  RentExtended, ExchangeITAvailMths,
  RentHoldOver, RentIncrement, UnitLRF
)
SELECT
  Units.UnitID,
  Units.GroupREF AS UnitGroupID,
  SchGrp.BLCD AS OrigBLCD,
  SchGrp.LeaseTermPrimary AS Term,
  DateAdd(
    "m", [SchGrp].[LeaseTermPrimary],
    [SchGrp].[BLCD]
  )-1 AS ProfileEOLDate,
  "Primary Lease" AS ProfileType,
  Units.UNITRENT AS RentLO,
  Units.UNITPDRENT AS RentPerdiem,
  Units.UNITSRENT AS RentSynd,
  Units.UNITSRENT AS RentActive,
  Units.UnitSynResid AS ResidualSynd,
  Units.UnitSalesResid,
  Banks.BankId,
  IIf(
    Nz([MTMRate], 0)= 0,
    [UnitSRent],
    CCur([Totalcost] * [MTMRate])
  ) AS RentMTM,
  SchGrp.LeaseTermNotice AS NoticeReqMonths,
  UnitOEC2.TotalCost AS UnitOEC,
  CCur(
    [TotalCost] * [LeaseExtOption1LRF]
  ) AS RentExtended,
  SchGrp.ExchangeITMonth AS ExchangeITAvailMths,
  Units.UNITSRENT AS RentHoldOver,
  'Monthly' AS RentIncrement,
  IIf(
    [UnitsRent] = 0
    Or [TotalCost] = 0,
    0,
    Round([UnitsRent] / [TotalCost], 6)
  ) AS UnitLRF
FROM
  (
    (
      (
        (
          Units
          INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID
        )
        INNER JOIN Banks ON SchGrp.Assignee_Bank = Banks.BankShortName
      )
      LEFT JOIN UnitExitOption ON Units.UnitID = UnitExitOption.UnitId
    )
    INNER JOIN vw_SixKeys ON Units.UnitID = vw_SixKeys.UnitID
  )
  INNER JOIN UnitOEC2 ON Units.UnitID = UnitOEC2.AssetId
WHERE
  (
    (
      (vw_SixKeys.MLOrig)= "FA"
    )
  );
