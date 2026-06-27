SELECT
  vw_SixKeys.ClientShortName AS Client,
  vw_SixKeys.FAID,
  [vw_SixKeys].[MLNo] & " | " & [vw_SixKeys].[Schedule] & " | " & [vw_Sixkeys].[UnitGroup] AS [MLA# | Sch | Grp],
  "(" & UCase([SchGrp].[Assigned_Indicator])& ") " & [vw_SixKeys].[BankShortName] AS Assignee,
  SchGrp.AssignmentDate AS [Date],
  SchGrp.BLCD,
  SchGrp.LeaseTermPrimary AS Term,
  SchGrp.SGprAllow1 AS MilesAllowed,
  SchGrp.LoanDate AS [Group Loan Date],
  SchGrp.SynMemo AS [Group Synd Memo],
  SchGrp.SGrpFSLSch AS [Group Service Fee],
  SchGrp.PVRate AS [Rate Used to Calc PV Service Fees],
  "" AS [SYND/FINAL ->],
  "" AS [Economics Detail ->],
  SchGrp.BLCD_Forecast AS BLCD,
  SchGrp.SaleDate AS [Sale Date],
  SchGrp.FundDate AS [Fund Date],
  SchGrp.SynFee AS [Synd Fee],
  SchGrp.BrokerFee AS [Broker Fee],
  SchGrp.SalePrice AS [Sale Price],
  SchGrp.SgrpSynPenalty AS [Group Term Penalty],
  IIf(
    [SchGrp].[SalePrice] > 0,
    (
      [SchGrp].[SalePrice] - [SchGrp].[OrigEquipCost] + SchGrp.[sgrpsynpenalty]
    ),
    0
  ) AS [Net Synd Penalty],
  SchGrp.PerDiemActual AS [PD/U],
  SchGrp.SgrpSynRebate AS [UTA/U],
  CCur(
    Nz(SchGrp.[SGrpFSLPmt], 0)+ Nz(SchGrp.[SGrpRSLTPmt], 0)
  ) AS [FMSS Fee/U],
  IIf(
    Trim(
      Nz(SchGrp.SyndMiscLabel1, "")& Nz(SchGrp.SyndMiscValue1)
    )= "",
    "",
    SchGrp.SyndMiscLabel1 & ": " & SchGrp.SyndMiscValue1
  ) AS [Misc 1],
  IIf(
    Trim(
      Nz(SchGrp.SyndMiscLabel2, "")& Nz(SchGrp.SyndMiscValue2)
    )= "",
    "",
    SchGrp.SyndMiscLabel2 & ": " & SchGrp.SyndMiscValue2
  ) AS [Misc 2],
  "" AS [Index ->],
  SchGrp.SGrpIndex AS [Index],
  SchGrp.SGrpIndDt AS [Index Date],
  SchGrp.SGprIndYr AS [Term/Yr],
  SchGrp.SGrpBonus AS Bonus,
  SchGrp.IndexRate AS Rate,
  SchGrp.SgrpYield AS Yield,
  [SchGrp].[sgrpyield] - [SchGrp].[IndexRate] AS Spread,
  "" AS [Rate Adjustment ->],
  SchGrp.BPChgLO AS [BP Change of],
  SchGrp.BPChgLO_Up AS Upward,
  SchGrp.BPChgLO_Down AS Downward,
  "" AS [LO/TARGET -->],
  "" AS [Economics Detail -->],
  SchGrp.BLCD_Projected AS [Proj BLCD],
  SchGrp.BLCDSetBy AS [Proj Set By],
  SchGrp.FundDate_Projected AS [Proj Fund Date],
  SchGrp.Fee_Projected AS [Proj Synd Fee],
  SchGrp.SgrpTermPenalty AS [Proj Group Term Penalty],
  Nz([SchGrp].[Fee_Projected], 0)+ Nz([SchGrp].[sgrptermPenalty], 0) AS [Proj Net Synd Penalty],
  SchGrp.PerDiemProjected AS [Proj PD/U],
  SchGrp.RebateUnitProjected AS [Proj UTA/U],
  SchGrp.SalesFSL AS [Proj FMSS Fee/U],
  "" AS [Index -->],
  SchGrp.SGrpSalesIndex AS [Proj Index],
  SchGrp.SGrpSalesIndDt AS [Proj Date],
  SchGrp.SGprSalesIndYr AS [Proj Term/Yr],
  SchGrp.SGrpSalesBonus AS [Proj Bonus],
  SchGrp.SgrpSalesIndRt AS [Proj Rate],
  SchGrp.SgrpSalesYield AS [Proj Yield],
  "" AS [UNIT -->],
  vw_SixKeys.AssetID,
  vw_SixKeys.PrintedDocs AS [Long Unit Description],
  UnitTotalOEC.UnitCost AS [Unit Cost],
  UnitTotalOEC.ChildPartCost AS [Child Part Cost],
  UnitTotalOEC.UnitTotalCost AS [Unit OEC],
  Units.UnitSRent AS [Sch Rent],
  Units.UnitSalesResid AS [Sch RV],
  Units.UnitRent AS [Synd Rent],
  Units.UnitSynResid AS [Synd RV],
  UnitTotalOEC.UnitLRF AS LRF,
  "" AS [Not Needed? -->],
  SchGrp.LeasePaidAdvArr,
  SchGrp.LeasePymtFreq,
  SchGrp.WithHeldPymts AS HoldBack,
  Units.UNITPDRENT AS PDRent,
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID
FROM
  (
    (
      (
        vw_SixKeys
        INNER JOIN UnitTotalOEC ON vw_SixKeys.AssetID = UnitTotalOEC.AssetID
      )
      INNER JOIN SchGrp ON vw_SixKeys.SGrpID = SchGrp.SGrpID
    )
    INNER JOIN Units ON vw_SixKeys.UnitID = Units.UnitID
  )
  INNER JOIN Schedule ON vw_SixKeys.SchID = Schedule.SchID
ORDER BY
  vw_SixKeys.SchID,
  vw_SixKeys.SGrpID;
