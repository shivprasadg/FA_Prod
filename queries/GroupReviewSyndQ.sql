SELECT
  DISTINCT Clients.clientCompanyName AS Client,
  Clients.ClientShNm AS ClientShortName,
  MstrLease.MLNo AS MasterLease,
  Clients_1.clientCompanyName AS Guarantor,
  Clients.clientCompanyName,
  Schedule.SchNo AS Schedule,
  Schedule.SchID,
  SchGrp.SGrpID AS GroupID,
  SchGrp.SGrpNO,
  SchGrp.SGrpNO AS GroupNum,
  SchGrp.UnitsInGroup,
  Count([Units].[UnitId])* 1 AS UnitCount,
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
  SchGrp.SalesFSL,
  Schedule.SchShortDesc AS ScheduleDesc,
  Units.UnitSynResid AS UnitResidual,
  Avg(Units.unitsalesresid) AS SalesUnitResidual,
  IIf(
    [UnitSynResid] > 0,
    Round(
      [UnitSynResid] / [OrigEquipCost],
      3
    ),
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
  CCur([unitrent]) AS UnitRentPayment,
  UnitTotalOEC.UnitLRF,
  Schedule.SchFSL,
  FSL.FSLYR1,
  SchGrp.SaleDate,
  SchGrp.SgrpTermPenalty,
  SchGrp.SgrpSynPenalty,
  SchGrp.WithHeldPymts,
  SyndGroupStripsUTA_Summed.UTAPer,
  SchGrp.SgrpSynRebate AS RebateUnitActual,
  SchGrp.SgrpSynPenalty AS SyndPenalty,
  SyndGroupStripsUTA_Summed.GroupStrips AS StripsNum,
  SyndGroupStripsUTA_Summed.GroupStripAmount AS StripsAmount,
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
    Nz(forms!groupform![PVRate], 0.05)
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
  InterimFunding_Days_Rate_Group.InterimIntRate,
  InterimFunding_Days_Rate_Group.AvgInterimLoanDays AS InterimLoanDays,
  InterimFunding_Days_Rate_Group.GroupInterimExp AS InterimFundingExpense,
  FAUserFullName() AS FullUserName,
  PerDiem_Days_Rate_Group.AvgPerDiem,
  PerDiem_Days_Rate_Group.PDDays,
  IIf(
    [SchGrp].[PerDiemActual] > 0,
    [SchGrp].[PerDiemActual] * [UnitsInGroup],
    IIf(
      [PerDiemGroup] > 0, [PerDiemGroup],
      [PerDiemProj]
    )
  ) AS PerDiemToDisplay,
  CDbl(
    Round(
      [LeasePymtForGroup] / [OrigEquipCost],
      7
    )
  ) AS LRF,
  IIf(
    [AllowedMilesPeriod] = 'P/Year',
    [SGprAllow1],
    (
      [SGprAllow1] /(
        [LeaseTermPrimary] - CInt(
          Nz([LeaseTermNotice], 0)
        )
      )* 12
    )
  ) AS MPY,
  CCur([OrigEquipCost] / [unitsingroup]) AS CostPerUnit,
  IIf(
    [SalePrice] > 0,
    (
      ([SalePrice] / [UnitsInGroup])- [UnitTotalCost]
    ),
    0
  ) AS FeeUnit,
  [FeeUnit] * [UnitCount] AS FeeUnitCount,
  CCur([SalePrice])-([Forms]![GroupForm]![CalcOEC]) AS FeeGroup,
  SchGrp.LoanDate,
  Sum(Units.UnitSynResid) AS GroupResidual,
  SchGrp.origEquipCost AS TotalCost,
  UnitTotalOEC.UnitTotalCost,
  SchGrp.LeasePymtForGroup,
  CCur([UnitRent]) AS LeasePymtUnit,
  SchGrp.SGrpNO AS [Group],
  CCur(
    Nz([HoldbackConversion], 0)
  ) AS HBConversion,
  CCur(
    Nz([HoldbackResidual], 0)
  ) AS HBResidual,
  CDbl(
    Nz([WithHeldPymts], 0)
  ) AS WithHeldNum,
  [LeasePymtForGroup] * [WithHeldPymts] AS WithHeldAmount,
  SchGrp.LastUpdatedDate,
  SchGrp.LastUpdatedBy,
  SchGrp.ScheduleGroupPaymentTerm,
  SchGrp.SGrpFSLPmt,
  SchGrp.SGrpRSLTPmt,
  Schedule.SchShortDesc,
  SchGrp.SyndMiscLabel1,
  SchGrp.SyndMiscValue1,
  SchGrp.SyndMiscLabel2,
  SchGrp.SyndMiscValue2,
  First(SchGrp.SynMemo) AS SynMemo,
  Units.UnitSRent AS SalesRent,
  SchGrp.RebateUnitProjected,
  Units.UnitRent AS Rent,
  UnitLevelTaxes.FeeTaxable AS AncillaryFee,
  CCur(
    Nz([UnitTaxes], 0)
  ) AS UnitTax,
  IIf(
    [UnitTaxes] > 0, [UnitlevelTaxes].[LState],
    "-"
  ) AS LocState,
  IIf(
    [UnitTaxes] > 0, [UnitlevelTaxes].[LTaxRate],
    0
  ) AS LTaxRate,
  SchGrp.SgrpDesc AS GroupDesc
FROM
  (
    (
      (
        (
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
            LEFT JOIN Banks ON SchGrp.Assignee_Bank = Banks.BankShortName
          )
          INNER JOIN UnitTotalOEC ON Units.UnitID = UnitTotalOEC.AssetID
        )
        INNER JOIN SyndGroupStripsUTA_Summed ON SchGrp.SGrpID = SyndGroupStripsUTA_Summed.GroupID
      )
      INNER JOIN InterimFunding_Days_Rate_Group ON SchGrp.SGrpID = InterimFunding_Days_Rate_Group.GroupREF
    )
    INNER JOIN PerDiem_Days_Rate_Group ON SchGrp.SGrpID = PerDiem_Days_Rate_Group.GroupREF
  )
  INNER JOIN UnitLevelTaxes ON Units.UnitID = UnitLevelTaxes.UnitID
GROUP BY
  Clients.ClientShNm,
  MstrLease.MLNo,
  Clients_1.clientCompanyName,
  Clients.clientCompanyName,
  Schedule.SchNo,
  Schedule.SchID,
  SchGrp.SGrpID,
  SchGrp.UnitsInGroup,
  SchGrp.OrigEquipCost,
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary,
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
  Units.UnitSynResid,
  IIf(
    [UnitSynResid] > 0,
    Round(
      [UnitSynResid] / [OrigEquipCost],
      3
    ),
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
  CCur([unitrent]),
  UnitTotalOEC.UnitLRF,
  Schedule.SchFSL,
  FSL.FSLYR1,
  SchGrp.SaleDate,
  SchGrp.SgrpTermPenalty,
  SchGrp.WithHeldPymts,
  SyndGroupStripsUTA_Summed.UTAPer,
  SchGrp.SgrpSynRebate,
  SchGrp.SgrpSynPenalty,
  SyndGroupStripsUTA_Summed.GroupStrips,
  SyndGroupStripsUTA_Summed.GroupStripAmount,
  CCur(
    Nz(
      [SchGrp].[PerDiemActual] * [unitsingroup],
      0
    )
  ),
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
  InterimFunding_Days_Rate_Group.InterimIntRate,
  InterimFunding_Days_Rate_Group.AvgInterimLoanDays,
  InterimFunding_Days_Rate_Group.GroupInterimExp,
  FAUserFullName(),
  PerDiem_Days_Rate_Group.AvgPerDiem,
  PerDiem_Days_Rate_Group.PDDays,
  IIf(
    [SchGrp].[PerDiemActual] > 0,
    [SchGrp].[PerDiemActual] * [UnitsInGroup],
    IIf(
      [PerDiemGroup] > 0, [PerDiemGroup],
      [PerDiemProj]
    )
  ),
  CDbl(
    Round(
      [LeasePymtForGroup] / [OrigEquipCost],
      7
    )
  ),
  IIf(
    [AllowedMilesPeriod] = 'P/Year',
    [SGprAllow1],
    (
      [SGprAllow1] /(
        [LeaseTermPrimary] - CInt(
          Nz([LeaseTermNotice], 0)
        )
      )* 12
    )
  ),
  CCur([OrigEquipCost] / [unitsingroup]),
  SchGrp.LoanDate,
  SchGrp.origEquipCost,
  UnitTotalOEC.UnitTotalCost,
  SchGrp.LeasePymtForGroup,
  CCur([UnitRent]),
  SchGrp.SGrpNO,
  CCur(
    Nz([HoldbackConversion], 0)
  ),
  CCur(
    Nz([HoldbackResidual], 0)
  ),
  CDbl(
    Nz([WithHeldPymts], 0)
  ),
  [LeasePymtForGroup] * [WithHeldPymts],
  SchGrp.LastUpdatedDate,
  SchGrp.LastUpdatedBy,
  SchGrp.ScheduleGroupPaymentTerm,
  SchGrp.SGrpFSLPmt,
  SchGrp.SGrpRSLTPmt,
  Schedule.SchShortDesc,
  SchGrp.SyndMiscLabel1,
  SchGrp.SyndMiscValue1,
  SchGrp.SyndMiscLabel2,
  SchGrp.SyndMiscValue2,
  Units.UnitSRent,
  SchGrp.RebateUnitProjected,
  Units.UnitRent,
  UnitLevelTaxes.FeeTaxable,
  CCur(
    Nz([UnitTaxes], 0)
  ),
  IIf(
    [UnitTaxes] > 0, [UnitlevelTaxes].[LState],
    "-"
  ),
  IIf(
    [UnitTaxes] > 0, [UnitlevelTaxes].[LTaxRate],
    0
  ),
  SchGrp.SgrpDesc,
  SchGrp.SGrpNO,
  Units.UnitSynResid,
  Schedule.SchFSL,
  SchGrp.SgrpSynPenalty,
  SchGrp.SGrpFSLPmt,
  SchGrp.SGrpNO,
  Schedule.SchShortDesc
HAVING
  (
    (
      (SchGrp.SGrpID)= [forms]![groupform]![sgrpid]
    )
  );
