SELECT
  DISTINCTROW EXITOptionsGAPReport.AssetID,
  2 AS OptionTypeID,
  IIf([Early] = "Yes",-1, 0) AS TerminateEarly,
  EXITOptionsGAPReport.ProjRV AS Residual,
  EXITOptionsGAPReport.ProjTermPenalty AS TerminateFee,
  EXITOptionsGAPReport.note AS Notes,
  EXITOptionsGAPReport.OffLeaseDate AS EndOrActualDate INTO ExitTermTemp
FROM
  EXITOptions_WorkTable
  INNER JOIN EXITOptionsGAPReport ON EXITOptions_WorkTable.UnitID = EXITOptionsGAPReport.AssetID;
