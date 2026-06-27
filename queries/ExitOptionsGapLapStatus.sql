SELECT
  EXITOptionsGAPReport.AssetID,
  IIf(
    IsNull([EarlyGapMsg]),
    Null,
    [EarlyGapMsg]
  )& " " & IIf(
    IsNull([OverLapMsg]),
    Null,
    [OverLapMsg]
  ) AS GapLapMsg,
  EXITOptionsGAPReport.Early,
  EXITOptionsGAPReport.GAP,
  EXITOptionsGAPReport.MonthsDiff
FROM
  EXITOptionsGAPReport;
