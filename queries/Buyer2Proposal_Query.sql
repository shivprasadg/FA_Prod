SELECT
  RemarketingReportAllUnits.UnitId,
  Nz([UnitSubTypeAbbr], '')& ' - ' & [UnitYr] & ' ' & [UnitMake] & ' ' & [UnitModel] & " " & [Axle] AS [Desc],
  RemarketingReportAllUnits.UnitStatus AS STS,
  BuyersProposalUnits.ProposalID,
  IIf(
    IsNull([InspectionLinkWeb]),
    "No",
    "Yes"
  ) AS Insp,
  RemarketingReportAllUnits.InspectionLinkWeb,
  Format([TargetSalePrice], "$#,###") AS Target,
  RemarketingReportAllUnits.ODORead AS ODO,
  RemarketingReportAllUnits.ODOSource,
  RemarketingReportAllUnits.DaysInInventory AS DaysIn,
  RemarketingReportAllUnits.SourceCompany,
  RemarketingReportAllUnits.Schedule,
  [LCity] & ", " & [LState] & ". " & [LZip] AS Location,
  RemarketingReportAllUnits.LocationID,
  RemarketingReportAllUnits.Latitude,
  RemarketingReportAllUnits.Longitude
FROM
  (
    BuyersProposalUnits
    INNER JOIN RemarketingReportAllUnits ON BuyersProposalUnits.UnitID = RemarketingReportAllUnits.UnitID
  )
  LEFT JOIN LookUp_UnitSubTypes ON RemarketingReportAllUnits.UnitSubType = LookUp_UnitSubTypes.UnitSubType;
