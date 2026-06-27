SELECT
  RemarketingReportAllUnits.UnitID,
  RemarketingReportAllUnits.SourceCompany,
  RemarketingReportAllUnits.[Sch|Grp],
  RemarketingReportAllUnits.UnitVIN,
  RemarketingReportAllUnits.UnitDescShortVer,
  RemarketingReportAllUnits.ODORead,
  RemarketingReportAllUnits.ODOReadDate,
  RemarketingReportAllUnits.UnitStatus,
  RemarketingReportAllUnits.DaysInInventory,
  RemarketingReportAllUnits.ParkDateProjected,
  RemarketingReportAllUnits.OffLeaseDate,
  RemarketingReportAllUnits.InvoiceSent,
  RemarketingReportAllUnits.Parked_v_EOL_Days,
  RemarketingReportAllUnits.Residual,
  RemarketingReportAllUnits.TermPenalty,
  RemarketingReportAllUnits.BuyOutDate,
  RemarketingReportAllUnits.BuyOut,
  RemarketingReportAllUnits.FMVBase,
  RemarketingReportAllUnits.FMV,
  RemarketingReportAllUnits.TargetSalePrice,
  RemarketingReportAllUnits.IFDate,
  RemarketingReportAllUnits.IFAmount,
  RemarketingReportAllUnits.IFBank
FROM
  RemarketingReportAllUnits
WHERE
  (
    (
      (
        RemarketingReportAllUnits.UnitStatus
      )= "R"
      Or (
        RemarketingReportAllUnits.UnitStatus
      )= "PS"
    )
  );
