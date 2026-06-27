SELECT
  DISTINCTROW Units.GroupREF,
  Units.UnitUnitNum,
  UnitExitOption.LeaseOptionID,
  UnitExitOption.UnitID,
  UnitExitOption.OptionTypeID,
  UnitExitOption.OptionTypeDesc,
  UnitExitOption.OptionSubType,
  UnitExitOption.StartOrPurposedDate AS [Start-Purposed],
  UnitExitOption.EndOrActualDate AS [End-Actual],
  UnitExitOption.TermMonths,
  UnitExitOption.Rent,
  UnitExitOption.Fees,
  UnitExitOption.Residual,
  UnitExitOption.ReplacedByScheduleID,
  UnitExitOption.ReplacedByGroupID,
  UnitExitOption.ReplacedByUnitID,
  UnitExitOption.TerminateDate,
  UnitExitOption.TerminateFee,
  UnitExitOption.FinalRentAmt,
  UnitExitOption.FinalRentDate,
  UnitExitOption.AssigneeNotified,
  UnitExitOption.AssigneeReplied,
  UnitExitOption.OrginalBLCD,
  UnitExitOption.OrigLeaseTerm,
  UnitExitOption.OrginalEOLDate,
  UnitExitOption.OrigResidual,
  UnitExitOption.TerminateEarly,
  UnitExitOption.UserName,
  UnitExitOption.DatePosted,
  UnitExitOption.Notes,
  UnitExitOption.UIDX,
  IIf(
    Nz([ReplacedByScheduleID], 0)= 0,
    Null,
    DLookUp(
      "SchNo", "Schedule", "SchId=" & [ReplacedByScheduleID] & ""
    )
  ) AS NewSchedule,
  IIf(
    Nz([ReplacedByGroupID], 0)= 0,
    Null,
    DLookUp(
      "sGrpNo", "SchGrp", "SgrpID=" & [ReplacedByGroupID] & ""
    )
  ) AS NewGroup,
  IIf(
    [OptionTypeID] = 1,
    Nz([End-Actual], [Start-Purposed]),
    Null
  ) AS ExchangeITDate,
  IIf(
    [OptionTypeID] <> 1,
    Null,
    IIf(
      IsNull([End-Actual]),
      "P",
      "A"
    )
  ) AS ExchangeITStatus,
  IIf(
    IsNull([NewSchedule]),
    Null,
    [NewSchedule] & "|" & [NewGroup] & IIf(
      IsNull([ReplacedByUnitID]),
      "|----",
      "|" & [ReplacedByUnitID]
    )
  ) AS ExchangeITSchGrpUnit,
  IIf(
    [OptionTypeID] > 2, [StartOrPurposedDate],
    Null
  ) AS ExtendITStartDate,
  IIf(
    [OptionTypeID] > 2, [EndOrActualDate],
    Null
  ) AS ExtendITEndDate,
  IIf([OptionTypeID] > 2, [Fees], Null) AS ExtendedFees,
  IIf([OptionTypeID] > 2, [Rent], Null) AS ExtendedRent,
  IIf(
    [OptionTypeID] = 2,
    Nz([End-Actual], [Start-Purposed]),
    Null
  ) AS TermDate,
  IIf(
    [OptionTypeID] <> 2,
    Null,
    IIf(
      IsNull([End-Actual]),
      "P",
      "A"
    )
  ) AS TermStatus,
  IIf(
    [OptionTypeID] <> 2,
    Null,
    Format([TerminateFee], "$#,###.##")
  ) AS GainLoss,
  IIf(
    [OptionTypeID] <> 2, Null, [TerminateEarly]
  ) AS TermEarly
FROM
  (
    UnitExitOption
    INNER JOIN Units ON UnitExitOption.UnitID = Units.UnitID
  )
  INNER JOIN EXITOptions_WorkTable ON UnitExitOption.UnitID = EXITOptions_WorkTable.UnitID
ORDER BY
  UnitExitOption.UnitID,
  UnitExitOption.OptionTypeDesc,
  UnitExitOption.StartOrPurposedDate DESC;
