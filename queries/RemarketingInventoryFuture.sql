SELECT
  RemarketingReportAllUnits.UnitID,
  RemarketingReportAllUnits.SourceCompany,
  RemarketingReportAllUnits.[Sch|Grp],
  RemarketingReportAllUnits.UnitVIN,
  RemarketingReportAllUnits.UnitDescShortVer,
  RemarketingReportAllUnits.ODORead,
  RemarketingReportAllUnits.ODOReadDate,
  RemarketingReportAllUnits.UnitStatus,
  RemarketingReportAllUnits.ParkDateProjected,
  RemarketingReportAllUnits.OffLeaseDate,
  RemarketingReportAllUnits.Residual,
  RemarketingReportAllUnits.TermPenalty,
  RemarketingReportAllUnits.BuyOutDate,
  RemarketingReportAllUnits.BuyOut,
  RemarketingReportAllUnits.FMVBase,
  RemarketingReportAllUnits.FMV,
  RemarketingReportAllUnits.TargetSalePrice
FROM
  RemarketingReportAllUnits
WHERE
  (
    (
      (
        RemarketingReportAllUnits.UnitStatus
      )<> "S"
      And (
        RemarketingReportAllUnits.UnitStatus
      )<> "C"
      And (
        RemarketingReportAllUnits.UnitStatus
      )<> 'SB'
      And (
        RemarketingReportAllUnits.UnitStatus
      )<> "KS"
      And (
        RemarketingReportAllUnits.UnitStatus
      )<> "D"
    )
  );
