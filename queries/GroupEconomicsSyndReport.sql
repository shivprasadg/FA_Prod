SELECT
  DISTINCT CDbl(
    [Forms]![GroupForm]![InterimLoanRate]
  ) AS IntrimIntRate,
  Max(Units.UnitIFDate) AS UnitIntrimDate,
  DateDiff(
    "d", [UnitIntrimDate], [FundDate]
  ) AS IntrimLoanDays,
  IIf(
    [IntrimIntRate] > 0,
    CDbl(
      IIf(
        IsDate([UnitIntrimDate]),
        [OrigEquipCost] + FV(
          ([IntrimIntRate] / 360),
          [IntrimLoanDays],
          0,
          [OrigEquipCost],
          1
        ),
        0
      )
    ),
    0
  ) AS IntrimFundingExpense,
  Clients.clientCompanyName AS Client,
  Clients.ClientShNm AS ClientShortName,
  MstrLease.MLNo AS MasterLease,
  Clients.clientCompanyName AS Guarantor,
  Schedule.SchNo AS Schedule,
  SchGrp.SGrpID AS GroupID,
  SchGrp.SGrpNO,
  SchGrp.SGrpNO AS GroupNum,
  SchGrp.UnitsInGroup,
  SchGrp.OrigEquipCost,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtForGroup,
  SchGrp.Assignee_Bank,
  Banks.BankName AS Assignee,
  SchGrp.Assigned_Indicator,
  SchGrp.AssignmentDate,
  SchGrp.FundDate,
  SchGrp.SalePrice,
  SchGrp.SGrpFSLPmt,
  SchGrp.SGrpFSLSch,
  SchGrp.LeasePaidAdvArr,
  SchGrp.LeasePymtFreq,
  SchGrp.ExchangeITGroup,
  SchGrp.ExchangeITMonth,
  SchGrp.ExchangeITType,
  SchGrp.SGrpBonus AS Bonus,
  SchGrp.SGrpIndex AS [Index],
  SchGrp.SGrpIndDt AS IndexDate,
  SchGrp.SGprIndYr AS IndexYear,
  SchGrp.IndexRate,
  SchGrp.SgrpYield AS Yield,
  SchGrp.EarlyTermNote,
  Count(Units.UnitID) AS UnitCount,
  SchGrp.SalesFSL,
  Units.UnitTax,
  SchGrp.RebateUnitProjected,
  Schedule.SchShortDesc AS ScheduleDesc,
  Units.UnitSynResid AS UnitResidual,
  Avg(Units.unitsalesresid) AS SalesUnitResidual,
  IIf(
    [UnitSynResid] > 0, [UnitSynResid] / [OrigEquipCost],
    0
  ) AS UnitRVPct,
  SchGrp.Fee_Projected,
  SchGrp.SGrpSalesBonus AS SalesBonus,
  SchGrp.SGrpSalesIndex AS SalesIndex,
  SchGrp.SGrpSalesIndDt AS SalesIndexDate,
  SchGrp.SGprSalesIndYr AS SalesIndexYear,
  SchGrp.SgrpSalesIndRt AS SalesIndexRate,
  SchGrp.SgrpSalesYield AS SalesYield,
  [sgrpyield] - [IndexRate] AS Spread,
  (
    [unitrent] + IIf(
      [schfsl] = No, 0, [sgrpfslpmt] - [fslyr1]
    )
  )* [UnitsInGroup] AS RentPmt,
  Schedule.SchFSL,
  FSL.FSLYR1,
  SchGrp.SaleDate,
  SchGrp.SgrpTermPenalty,
  SchGrp.SgrpSynPenalty,
  Units.UNITRENT AS Rent,
  Units.unitsrent AS SalesRent,
  SchGrp.WithHeldPymts,
  SchGrp.SgrpSynRebate AS RebateUnitActual,
  SchGrp.SgrpSynPenalty AS SyndPenalty,
  DateDiff("m", [BLCD], [AssignmentDate]) AS StripsNum,
  CCur(
    Nz(
      (
        (
          DateDiff("m", [BLCD], [AssignmentDate])
        )
      )*([SalesRent] * [unitsingroup]),
      0
    )
  ) AS StripsAmount,
  CCur(
    Nz(
      [SchGrp].[PerDiemActual] * [unitsingroup]
    )
  ) AS PerDiemActual,
  SchGrp.SGrpFSLPmt AS ServiceFee,
  Switch(
    [LeasePymtFreq] = "M", "Monthly", [LeasePymtFreq] = "Q",
    "Quarterly", [LeasePymtFreq] = "A",
    "Annualy", [LeasePymtFreq] = "S",
    "Semi-Annuly"
  ) AS PayFreq,
  Switch(
    [LeasePaidAdvArr] = "Adv", "Advance",
    [LeasePaidAdvArr] = "Arr", "Arrears"
  ) AS AdvanceArrears,
  IIf(
    Nz([ScheduleGroupPaymentTerm], 0)= 0,
    [LeaseTermPrimary],
    [ScheduleGroupPaymentTerm]
  ) AS FeeTerm,
  (
    (
      IIf(
        Nz([ScheduleGroupPaymentTerm], 0)= 0,
        [LeaseTermPrimary],
        [ScheduleGroupPaymentTerm]
      )
    )* [SGrpFSLPmt]
  )* [Unitsingroup] AS FullServiceFee,
  CDbl(
    Nz(
      [forms]![groupform]![pvfeerate],
      0.05
    )
  ) AS PVFeeRate,
  IIf(
    [forms]![groupform]![PVFeesYN] = True
    And [forms]![groupform]![ServFeePassTroughYN] = False,
    Round(
      PV(
        CDbl(
          (
            Nz([PVFeeRate], 0.05)/ 12
          )
        ),
        [FeeTerm],
        ([SGrpFSLPmt] * [UnitsInGroup])*-1,
        0,
        1
      ),
      2
    ),
    0
  ) AS ServFeePV,
  FAUserFullName() AS FullUserName,
  IIf(
    [forms].[groupform]![CalcPerDiemYN] = True,
    Max(
      CDbl(
        Nz([PerdiemTotal], 0)
      )
    ),
    CCur(
      Nz(
        [SchGrp].[PerDiemActual] * [unitsingroup]
      )
    )
  ) AS PerdiemGroup,
  CDbl(
    Round(
      [LeasePymtForGroup] / [OrigEquipCost],
      7
    )
  ) AS LRF,
  IIf(
    [AllowedMilesPeriod] = 'P/TERM',
    [SGprAllow1] /(
      [LeaseTermPrimary] - Nz([LeaseTermNotice], 0)
    ),
    [SGprAllow1]
  ) AS MPY,
  CCur([OrigEquipCost] / [unitsingroup]) AS CostPerUnit,
  IIf(
    [SalePrice] > 0,
    ([SalePrice] - [OrigEquipCost]),
    0
  ) AS FeeGroup,
  CCur(
    IIf(
      [SalePrice] > 0,
      (
        ([SalePrice] - [OrigEquipCost])
      )/ [UnitsInGroup],
      0
    )
  ) AS FeeUnit,
  SchGrp.LoanDate,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] & " " & [Axle] AS UnitDesc,
  Sum(Units.UnitSynResid) AS GroupResidual,
  SchGrp.origEquipCost AS TotalCost,
  CCur(
    [LeasePymtForGroup] / [UnitsinGroup]
  ) AS LeasePymtUnit,
  SchGrp.HoldbackResidual,
  SchGrp.HoldbackConversion,
  Schedule.SchShortDesc,
  SchGrp.SgrpDesc,
  SchGrp.SynMemo
FROM
  (
    (
      (
        (
          (
            (
              Clients
              INNER JOIN MstrLease ON Clients.clientID = MstrLease.Client
            )
            INNER JOIN (
              Schedule
              LEFT JOIN FSL ON Schedule.FSLSCH = FSL.ID
            ) ON MstrLease.MLID = Schedule.SchMLRef
          )
          INNER JOIN SchGrp ON Schedule.SchID = SchGrp.SGrpSchID
        )
        INNER JOIN Units ON SchGrp.SGrpID = Units.GroupREF
      )
      LEFT JOIN UnitGroupPerdiemsSyndTotals ON SchGrp.SGrpID = UnitGroupPerdiemsSyndTotals.GroupID
    )
    LEFT JOIN Clients AS Clients_1 ON MstrLease.MLGuarantor = Clients_1.clientID
  )
  LEFT JOIN Banks ON SchGrp.Assignee_Bank = Banks.BankShortName
GROUP BY
  CDbl(
    [Forms]![GroupForm]![InterimLoanRate]
  ),
  Clients.ClientShNm,
  MstrLease.MLNo,
  Clients.clientCompanyName,
  Schedule.SchNo,
  SchGrp.SGrpID,
  SchGrp.SGrpNO,
  SchGrp.UnitsInGroup,
  SchGrp.OrigEquipCost,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtForGroup,
  SchGrp.Assignee_Bank,
  Banks.BankName,
  SchGrp.Assigned_Indicator,
  SchGrp.AssignmentDate,
  SchGrp.FundDate,
  SchGrp.SalePrice,
  SchGrp.SGrpFSLSch,
  SchGrp.LeasePaidAdvArr,
  SchGrp.LeasePymtFreq,
  SchGrp.ExchangeITGroup,
  SchGrp.ExchangeITMonth,
  SchGrp.ExchangeITType,
  SchGrp.SGrpBonus,
  SchGrp.SGrpIndex,
  SchGrp.SGrpIndDt,
  SchGrp.SGprIndYr,
  SchGrp.IndexRate,
  SchGrp.SgrpYield,
  SchGrp.EarlyTermNote,
  SchGrp.SalesFSL,
  Units.UnitTax,
  SchGrp.RebateUnitProjected,
  Units.UnitSynResid,
  IIf(
    [UnitSynResid] > 0, [UnitSynResid] / [OrigEquipCost],
    0
  ),
  SchGrp.Fee_Projected,
  SchGrp.SGrpSalesBonus,
  SchGrp.SGrpSalesIndex,
  SchGrp.SGrpSalesIndDt,
  SchGrp.SGprSalesIndYr,
  SchGrp.SgrpSalesIndRt,
  SchGrp.SgrpSalesYield,
  [sgrpyield] - [IndexRate],
  (
    [unitrent] + IIf(
      [schfsl] = No, 0, [sgrpfslpmt] - [fslyr1]
    )
  )* [UnitsInGroup],
  Schedule.SchFSL,
  FSL.FSLYR1,
  SchGrp.SaleDate,
  SchGrp.SgrpTermPenalty,
  Units.UNITRENT,
  Units.unitsrent,
  SchGrp.WithHeldPymts,
  SchGrp.SgrpSynRebate,
  SchGrp.SgrpSynPenalty,
  CCur(
    Nz(
      [SchGrp].[PerDiemActual] * [unitsingroup]
    )
  ),
  SchGrp.SGrpFSLPmt,
  Switch(
    [LeasePymtFreq] = "M", "Monthly", [LeasePymtFreq] = "Q",
    "Quarterly", [LeasePymtFreq] = "A",
    "Annualy", [LeasePymtFreq] = "S",
    "Semi-Annuly"
  ),
  IIf(
    Nz([ScheduleGroupPaymentTerm], 0)= 0,
    [LeaseTermPrimary],
    [ScheduleGroupPaymentTerm]
  ),
  FAUserFullName(),
  CDbl(
    Round(
      [LeasePymtForGroup] / [OrigEquipCost],
      7
    )
  ),
  IIf(
    [AllowedMilesPeriod] = 'P/TERM',
    [SGprAllow1] /(
      [LeaseTermPrimary] - Nz([LeaseTermNotice], 0)
    ),
    [SGprAllow1]
  ),
  CCur([OrigEquipCost] / [unitsingroup]),
  SchGrp.LoanDate,
  [UnitYr] & " " & [UnitMake] & " " & [UnitModel] & " " & [UnitSubtype] & " " & [UnitType] & " " & [Axle],
  SchGrp.origEquipCost,
  CCur(
    [LeasePymtForGroup] / [UnitsinGroup]
  ),
  SchGrp.HoldbackResidual,
  SchGrp.HoldbackConversion,
  Schedule.SchShortDesc,
  SchGrp.SgrpDesc,
  SchGrp.SynMemo,
  SchGrp.SGrpNO,
  Units.UnitSynResid,
  Schedule.SchFSL,
  SchGrp.SgrpSynPenalty,
  SchGrp.SGrpFSLPmt
HAVING
  (
    (
      (SchGrp.SGrpID)= [forms]![groupform]![sgrpid]
    )
    AND (
      (Units.unitsrent)> 0
    )
  );
