SELECT
  DISTINCTROW AllLeasedUnits_Cost.Schedule,
  AllLeasedUnits_Cost.Unitnum,
  AllLeasedUnits_Cost.VIN,
  AllLeasedUnits_Cost.UnitGroup AS SCH_Grp,
  AllLeasedUnits_Cost.UnitCost,
  AllLeasedUnits_Cost.OBCMake,
  AllLeasedUnits_Cost.OBCInvoice,
  AllLeasedUnits_Cost.OBCCost,
  AllLeasedUnits_Cost.OBCInstaller,
  AllLeasedUnits_Cost.OBCInstall,
  AllLeasedUnits_Cost.OBCInstall_Invoice,
  AllLeasedUnits_Cost.WarrantyCost,
  AllLeasedUnits_Cost.OtherCost,
  AllLeasedUnits_Cost.FeesTaxesCost,
  [UnitCost] + [obcinstall] + [WarrantyCost] + [OtherCost] + [FeesTaxesCost] AS TotalCost
FROM
  AllLeasedUnits_Cost
WHERE
  (
    (
      (AllLeasedUnits_Cost.SchID)= [Forms]![ScheduleControl]![Schid]
    )
  )
ORDER BY
  AllLeasedUnits_Cost.VIN;
