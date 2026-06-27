SELECT
  DISTINCTROW DLookUp(
    "UnitUnitNum", "Units", "UnitId=" & [EXITOptions_WorkTable].[UnitId] & ""
  ) AS UnitNum,
  DLookUp(
    "UnitVin", "Units", "UnitId=" & [EXITOptions_WorkTable].[UnitId] & ""
  ) AS VIN,
  DLookUp(
    "BLCD", "SchGrp", "SGrpID=" & [GroupId] & ""
  ) AS BLCD,
  DLookUp(
    "LeaseTermPrimary", "SchGrp", "SGrpID=" & [GroupId] & ""
  ) AS LeaseTerm,
  DateAdd("m", [LeaseTerm], [BLCD])-1 AS Termination,
  DateAdd(
    "m",
    Nz(
      DLookUp(
        "ExchangeITMonth", "SchGrp", "SGrpID=" & [GroupId] & ""
      ),
      0
    ),
    [BLCD]
  )-1 AS ExchangeITDate,
  ([Termination] - 91) AS NoticeDate,
  IIf(
    [ExchangeITDate] = "TBD",
    100,
    DateDiff(
      "m",
      Date(),
      [ExchangeITDate]
    )
  ) AS MonthsUnitEXIT,
  UnitExitOption.GroupId AS Expr1,
  UnitExitOption.OptionSelected AS Expr2,
  UnitExitOption.ExtensionStart AS Expr3,
  UnitExitOption.ExtensionTerm AS Expr4,
  UnitExitOption.PrevTermination AS Expr5,
  UnitExitOption.CurrTermination AS Expr6,
  UnitExitOption.ExchangeDate AS Expr7,
  UnitExitOption.ExchangedMatched AS Expr8,
  UnitExitOption.ReplaceScheduleId AS Expr9,
  UnitExitOption.ReplacedByUnitId AS Expr10,
  UnitExitOption.ExchangePool AS Expr11,
  UnitExitOption.DateSeleted AS Expr12,
  UnitExitOption.PrevBLCD AS Expr13,
  UnitExitOption.LFR AS Expr14,
  UnitExitOption.Rent AS Expr15,
  UnitExitOption.ClientNoticeSent AS Expr16,
  UnitExitOption.BankNoticeSent AS Expr17,
  UnitExitOption.NoticeConfrimed AS Expr18,
  UnitExitOption.AssigneeBankId,
  UnitExitOption.EstTurnInDate AS Expr19,
  UnitExitOption.ActualTurnInDate AS Expr20,
  UnitExitOption.OptionStatus AS Expr21
FROM
  EXITOptions_WorkTable
  LEFT JOIN UnitExitOption ON EXITOptions_WorkTable.UnitID = UnitExitOption.UnitId;
