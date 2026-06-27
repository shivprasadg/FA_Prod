SELECT
  ATF_Schedule_Group.ATFID,
  ATF_Schedule_Group.ATFSchName,
  ATF_Schedule_Group.ATFGroupSpecID,
  ATF_Schedule_Group.ATFGroupSpecDesc,
  ATF_Schedule_Group.ATFUnitSpecID,
  ATF_Schedule_Group.LeaseType,
  ATF_Schedule_Group.PrimaryLeaseTerm,
  ATF_Schedule_Group.PrimaryLeaseNotice,
  ATF_Schedule_Group.LeasePymtFreq,
  ATF_Schedule_Group.LeasePaidAdvArr,
  ATF_Schedule_Group.ExchangeITGroup,
  ATF_Schedule_Group.ExchangeITType,
  ATF_Schedule_Group.ExchangeITMonth,
  ATF_Schedule_Group.ExtendITOptionTerm,
  ATF_Schedule_Group.ExtendedRentAvg,
  ATF_Schedule_Group.MTMRentAvg,
  ATF_Schedule_Group.SwapIndex,
  ATF_Schedule_Group.SwapDate,
  ATF_Schedule_Group.SwapRate,
  ATF_Schedule_Group.SwapTermYears,
  ATF_Schedule_Group.SwapSpread,
  ATF_Schedule_Group.SwapYield,
  ATF_Schedule_Group.SwapBonus,
  ATF_Schedule_Group.SwapAdjustment,
  ATF_Schedule_Group.ResidualAvg,
  DLookUp(
    "StartsOn", "Lookup_LeaseStartsOn",
    "ID=" & [BLCDBasisID]
  ) AS NegBLCD,
  DLookUp(
    "StartsOn", "Lookup_LeaseStartsOn",
    "ID=" & [PerDiemBasisID]
  ) AS PerDiemType,
  ATF_UnitSpecCostRollup.UnitCost,
  ATF_UnitSpecCostRollup.PartCost,
  ATF_UnitSpecCostRollup.TotalUnitCost,
  CCur(
    Nz([UATPerUnit], 0)
  ) AS UAT,
  ATF_UnitSpecCostRollup.PerDiemRent,
  ATF_Schedule_Group.SchRent,
  ATF_UnitSpecCostRollup.LRFSch,
  ATF_Schedule_Group.FirstExtendedRent,
  ATF_UnitSpecCostRollup.LRFExt,
  ATF_Schedule_Group.MTMRent,
  ATF_UnitSpecCostRollup.LRFMTM,
  ATF_UnitSpecCostRollup.NumUnitsInSpec AS Units
FROM
  ATF_Schedule_Group
  LEFT JOIN ATF_UnitSpecCostRollup ON ATF_Schedule_Group.ATFGroupSpecID = ATF_UnitSpecCostRollup.ATFGroupSpecID
GROUP BY
  ATF_Schedule_Group.ATFID,
  ATF_Schedule_Group.ATFSchName,
  ATF_Schedule_Group.ATFGroupSpecID,
  ATF_Schedule_Group.ATFGroupSpecDesc,
  ATF_Schedule_Group.ATFUnitSpecID,
  ATF_Schedule_Group.LeaseType,
  ATF_Schedule_Group.PrimaryLeaseTerm,
  ATF_Schedule_Group.PrimaryLeaseNotice,
  ATF_Schedule_Group.LeasePymtFreq,
  ATF_Schedule_Group.LeasePaidAdvArr,
  ATF_Schedule_Group.ExchangeITGroup,
  ATF_Schedule_Group.ExchangeITType,
  ATF_Schedule_Group.ExchangeITMonth,
  ATF_Schedule_Group.ExtendITOptionTerm,
  ATF_Schedule_Group.ExtendedRentAvg,
  ATF_Schedule_Group.MTMRentAvg,
  ATF_Schedule_Group.SwapIndex,
  ATF_Schedule_Group.SwapDate,
  ATF_Schedule_Group.SwapRate,
  ATF_Schedule_Group.SwapTermYears,
  ATF_Schedule_Group.SwapSpread,
  ATF_Schedule_Group.SwapYield,
  ATF_Schedule_Group.SwapBonus,
  ATF_Schedule_Group.SwapAdjustment,
  ATF_Schedule_Group.ResidualAvg,
  DLookUp(
    "StartsOn", "Lookup_LeaseStartsOn",
    "ID=" & [BLCDBasisID]
  ),
  DLookUp(
    "StartsOn", "Lookup_LeaseStartsOn",
    "ID=" & [PerDiemBasisID]
  ),
  ATF_UnitSpecCostRollup.UnitCost,
  ATF_UnitSpecCostRollup.PartCost,
  ATF_UnitSpecCostRollup.TotalUnitCost,
  CCur(
    Nz([UATPerUnit], 0)
  ),
  ATF_UnitSpecCostRollup.PerDiemRent,
  ATF_Schedule_Group.SchRent,
  ATF_UnitSpecCostRollup.LRFSch,
  ATF_Schedule_Group.FirstExtendedRent,
  ATF_UnitSpecCostRollup.LRFExt,
  ATF_Schedule_Group.MTMRent,
  ATF_UnitSpecCostRollup.LRFMTM,
  ATF_UnitSpecCostRollup.NumUnitsInSpec;
