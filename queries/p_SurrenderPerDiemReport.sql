SELECT Six.ClientShortName As Client, Six.Schedule, Six.UnitGroup as [Group], Six.UnitID, Six.Unitnum, PDS.BuyOutCost, PDS.BuyOutDate, PDS.BuyDateEffective AS EffectiveBuyOut,InspectionPassed
,Case WHEN InspectionPassed < BuyDateEffective then 0 else DATEDIFF(dd, Coalesce(Nullif(PDS.BuyDateEffective,''),PDS.BuyOutDate), COALESCE (PDS.InspectionPassed, GETDATE()))  End AS Days
, PDS.PerDiemSurrender AS HoldOver , Case WHEN InspectionPassed < BuyDateEffective then 0 else PDS.PerDiemSurrender * DATEDIFF(dd, Coalesce(Nullif(PDS.BuyDateEffective,''),PDS.BuyOutDate), COALESCE (PDS.InspectionPassed, GETDATE())) END AS TotalHoldOver
, Six.OffLeaseActual AS ActualOffLease, UnitStatus
FROM vw_SixKeys AS Six INNER JOIN vw_PerDiemSurrender AS PDS ON Six.AssetID = PDS.UnitID 

