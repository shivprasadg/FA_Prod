SELECT
  DISTINCT Clients.clientCompanyName AS Client,
  Clients.ClientShNm AS ClientShortName,
  MstrLease.MLNo AS MasterLease,
  Clients.clientCompanyName AS Guarantor,
  Schedule.SchNo AS Schedule,
  SchGrp.SGrpID AS GroupID,
  SchGrp.SGrpNO,
  SchGrp.SGrpNO AS GroupNum,
  Nz(
    [SgrpDesc], "Units in Group: (" & [UnitsInGroup] & ")"
  ) AS GroupDesc,
  SchGrp.UnitsInGroup,
  SchGrp.OrigEquipCost,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtForGroup,
  SchGrp.Assignee_Bank,
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
  CCur(
    Nz([UnitTaxes], 0)
  ) AS UnitTax,
  SchGrp.RebateUnitProjected,
  Units.UnitSynResid AS UnitResidual,
  Avg(Units.unitsalesresid) AS SalesUnitResidual,
  SchGrp.Fee_Projected,
  SchGrp.SGrpSalesBonus AS SalesBonus,
  SchGrp.SGrpSalesIndex AS SalesIndex,
  SchGrp.SGrpSalesIndDt AS SalesIndexDate,
  SchGrp.SGprSalesIndYr AS SalesIndexYear,
  SchGrp.SgrpSalesIndRt AS SalesIndexRate,
  SchGrp.SgrpSalesYield AS SalesYield,
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
  SyndGroupStripsUTA_Summed.GroupStrips AS StripsNum,
  SyndGroupStripsUTA_Summed.GroupStripAmount AS StripsAmount,
  CCur(
    Nz(
      [SchGrp].[PerDiemActual] * [unitsingroup]
    )
  ) AS PerDiemActual,
  SchGrp.SGrpFSLPmt AS ServiceFee,
  UnitLevelTaxes.FeeTaxable AS AncillaryFee,
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
    )*(
      [SGrpFSLPmt] + Nz([AncillaryFee], 0)
    )
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
  forms.groupform![InterimLoanRate] AS IntrimIntRate,
  DateDiff("d", [LoanDate], [FundDate]) AS IntrimLoanDays,
  IIf(
    [forms].[groupform]![CalcIntrimYN] = True,
    CDbl(
      IIf(
        IsDate([LoanDate]),
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
  CCur(
    Nz([HoldbackConversion], 0)
  ) AS HBConversion,
  CCur(
    Nz([HoldbackResidual], 0)
  ) AS HBResidual,
  [LeasePymtForGroup] * [WithHeldPymts] AS WithHeldAmount,
  IIf(
    [UnitTaxes] > 0, [UnitlevelTaxes].[LocState],
    "-"
  ) AS LocState,
  IIf(
    [UnitTaxes] > 0, [UnitlevelTaxes].[LTaxRate],
    0
  ) AS LTaxRate,
  SchGrp.LastUpdatedBy,
  SchGrp.LastUpdatedDate,
  SchGrp.SGrpRSLTPmt
FROM
  (
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
    INNER JOIN UnitLevelTaxes ON Units.UnitID = UnitLevelTaxes.UnitID
  )
  INNER JOIN SyndGroupStripsUTA_Summed ON SchGrp.SGrpID = SyndGroupStripsUTA_Summed.GroupID
GROUP BY
  Clients.ClientShNm,
  MstrLease.MLNo,
  Clients.clientCompanyName,
  Schedule.SchNo,
  SchGrp.SGrpID,
  SchGrp.SGrpNO,
  Nz(
    [SgrpDesc], "Units in Group: (" & [UnitsInGroup] & ")"
  ),
  SchGrp.UnitsInGroup,
  SchGrp.OrigEquipCost,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
  SchGrp.LeasePymtForGroup,
  SchGrp.Assignee_Bank,
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
  CCur(
    Nz([UnitTaxes], 0)
  ),
  SchGrp.RebateUnitProjected,
  Units.UnitSynResid,
  SchGrp.Fee_Projected,
  SchGrp.SGrpSalesBonus,
  SchGrp.SGrpSalesIndex,
  SchGrp.SGrpSalesIndDt,
  SchGrp.SGprSalesIndYr,
  SchGrp.SgrpSalesIndRt,
  SchGrp.SgrpSalesYield,
  Schedule.SchFSL,
  FSL.FSLYR1,
  SchGrp.SaleDate,
  SchGrp.SgrpTermPenalty,
  Units.UNITRENT,
  Units.unitsrent,
  SchGrp.WithHeldPymts,
  SchGrp.SgrpSynRebate,
  SchGrp.SgrpSynPenalty,
  SyndGroupStripsUTA_Summed.GroupStrips,
  SyndGroupStripsUTA_Summed.GroupStripAmount,
  CCur(
    Nz(
      [SchGrp].[PerDiemActual] * [unitsingroup]
    )
  ),
  SchGrp.SGrpFSLPmt,
  UnitLevelTaxes.FeeTaxable,
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
  forms.groupform![InterimLoanRate],
  DateDiff("d", [LoanDate], [FundDate]),
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
  CCur(
    Nz([HoldbackConversion], 0)
  ),
  CCur(
    Nz([HoldbackResidual], 0)
  ),
  [LeasePymtForGroup] * [WithHeldPymts],
  IIf(
    [UnitTaxes] > 0, [UnitlevelTaxes].[LocState],
    "-"
  ),
  IIf(
    [UnitTaxes] > 0, [UnitlevelTaxes].[LTaxRate],
    0
  ),
  SchGrp.LastUpdatedBy,
  SchGrp.LastUpdatedDate,
  SchGrp.SGrpRSLTPmt,
  (
    [unitrent] + IIf(
      [schfsl] = No, 0, [sgrpfslpmt] - [fslyr1]
    )
  )* [UnitsInGroup],
  SchGrp.SGrpNO,
  Units.UnitSynResid,
  Schedule.SchFSL,
  SchGrp.SgrpSynPenalty,
  SchGrp.SGrpFSLPmt
HAVING
  (
    (
      (SchGrp.SGrpID)= forms!groupform!sgrpid
    )
    And (
      (Units.unitsrent)> 0
    )
  );
