SELECT
  GroupFinalizationReport_base.Client,
  GroupFinalizationReport_base.Schedule,
  GroupFinalizationReport_base.UnitGroup,
  GroupFinalizationReport_base.Assigned_Indicator AS Assigned,
  GroupFinalizationReport_base.AssignmentDate,
  IIf(
    [GroupFinalizationReport_base].[Fee_Projected] = 0,
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='Fee_Projected'"
    )
  ) AS SyndFeeProj,
  IIf(
    [GroupFinalizationReport_base].[SynFee] = 0,
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='SynFee'"
    )
  ) AS SynFee,
  IIf(
    IsNull(
      [GroupFinalizationReport_base].[SGrpSalesIndex]
    ),
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='SGrpSalesIndex'"
    )
  ) AS [Index(LO)],
  IIf(
    IsNull(
      [GroupFinalizationReport_base].[SGrpIndex]
    ),
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='SGrpIndex'"
    )
  ) AS [Index],
  IIf(
    IsNull(
      [GroupFinalizationReport_base].[SgrpSalesYield]
    ),
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='SgrpSalesYield'"
    )
  ) AS [Yield(LO)],
  IIf(
    IsNull(
      [GroupFinalizationReport_base].[SgrpYield]
    ),
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='SgrpYield'"
    )
  ) AS Yield,
  IIf(
    IsNull(
      [GroupFinalizationReport_base].[SGrpRALB]
    ),
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='SGrpRALB'"
    )
  ) AS [RAL Back],
  IIf(
    IsNull(
      [GroupFinalizationReport_base].[SchIn]
    ),
    DLookUp(
      "Person", "GroupFinalizedPersons",
      "Fieldname='SchIn'"
    )
  ) AS SchIn,
  GroupFinalizationReport_base.PDiemRent AS PDUnit,
  GroupFinalizationReport_base.BillToAddress,
  GroupFinalizationReport_base.BLCDBasis,
  GroupFinalizationReport_base.UnitSyndRent,
  GroupFinalizationReport_base.UnitResidual,
  GroupFinalizationReport_base.SchID,
  GroupFinalizationReport_base.SGrpID
FROM
  GroupFinalizationReport_base
ORDER BY
  GroupFinalizationReport_base.Client,
  GroupFinalizationReport_base.Schedule;
