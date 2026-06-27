Operation =1
Option =0
Begin InputTables
    Name ="vw_SixKeys"
    Name ="UnitTotalOEC"
    Name ="SchGrp"
    Name ="Units"
    Name ="Schedule"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.FAID"
    Alias ="MLA# | Sch | Grp"
    Expression ="[vw_SixKeys].[MLNo] & \" | \" & [vw_SixKeys].[Schedule] & \" | \" & [vw_Sixkeys]"
        ".[UnitGroup]"
    Alias ="Assignee"
    Expression ="\"(\" & UCase([SchGrp].[Assigned_Indicator]) & \") \" & [vw_SixKeys].[BankShortN"
        "ame]"
    Alias ="Date"
    Expression ="SchGrp.AssignmentDate"
    Expression ="SchGrp.BLCD"
    Alias ="Term"
    Expression ="SchGrp.LeaseTermPrimary"
    Alias ="MilesAllowed"
    Expression ="SchGrp.SGprAllow1"
    Alias ="Group Loan Date"
    Expression ="SchGrp.LoanDate"
    Alias ="Group Synd Memo"
    Expression ="SchGrp.SynMemo"
    Alias ="Group Service Fee"
    Expression ="SchGrp.SGrpFSLSch"
    Alias ="Rate Used to Calc PV Service Fees"
    Expression ="SchGrp.PVRate"
    Alias ="SYND/FINAL ->"
    Expression ="\"\""
    Alias ="Economics Detail ->"
    Expression ="\"\""
    Alias ="BLCD"
    Expression ="SchGrp.BLCD_Forecast"
    Alias ="Sale Date"
    Expression ="SchGrp.SaleDate"
    Alias ="Fund Date"
    Expression ="SchGrp.FundDate"
    Alias ="Synd Fee"
    Expression ="SchGrp.SynFee"
    Alias ="Broker Fee"
    Expression ="SchGrp.BrokerFee"
    Alias ="Sale Price"
    Expression ="SchGrp.SalePrice"
    Alias ="Group Term Penalty"
    Expression ="SchGrp.SgrpSynPenalty"
    Alias ="Net Synd Penalty"
    Expression ="IIf([SchGrp].[SalePrice]>0,([SchGrp].[SalePrice]-[SchGrp].[OrigEquipCost]+SchGrp"
        ".[sgrpsynpenalty]),0)"
    Alias ="PD/U"
    Expression ="SchGrp.PerDiemActual"
    Alias ="UTA/U"
    Expression ="SchGrp.SgrpSynRebate"
    Alias ="FMSS Fee/U"
    Expression ="CCur(Nz(SchGrp.[SGrpFSLPmt],0)+Nz(SchGrp.[SGrpRSLTPmt],0))"
    Alias ="Misc 1"
    Expression ="IIf(Trim(Nz(SchGrp.SyndMiscLabel1,\"\") & Nz(SchGrp.SyndMiscValue1))=\"\",\"\",S"
        "chGrp.SyndMiscLabel1 & \": \" & SchGrp.SyndMiscValue1)"
    Alias ="Misc 2"
    Expression ="IIf(Trim(Nz(SchGrp.SyndMiscLabel2,\"\") & Nz(SchGrp.SyndMiscValue2))=\"\",\"\",S"
        "chGrp.SyndMiscLabel2 & \": \" & SchGrp.SyndMiscValue2)"
    Alias ="Index ->"
    Expression ="\"\""
    Alias ="Index"
    Expression ="SchGrp.SGrpIndex"
    Alias ="Index Date"
    Expression ="SchGrp.SGrpIndDt"
    Alias ="Term/Yr"
    Expression ="SchGrp.SGprIndYr"
    Alias ="Bonus"
    Expression ="SchGrp.SGrpBonus"
    Alias ="Rate"
    Expression ="SchGrp.IndexRate"
    Alias ="Yield"
    Expression ="SchGrp.SgrpYield"
    Alias ="Spread"
    Expression ="[SchGrp].[sgrpyield]-[SchGrp].[IndexRate]"
    Alias ="Rate Adjustment ->"
    Expression ="\"\""
    Alias ="BP Change of"
    Expression ="SchGrp.BPChgLO"
    Alias ="Upward"
    Expression ="SchGrp.BPChgLO_Up"
    Alias ="Downward"
    Expression ="SchGrp.BPChgLO_Down"
    Alias ="LO/TARGET -->"
    Expression ="\"\""
    Alias ="Economics Detail -->"
    Expression ="\"\""
    Alias ="Proj BLCD"
    Expression ="SchGrp.BLCD_Projected"
    Alias ="Proj Set By"
    Expression ="SchGrp.BLCDSetBy"
    Alias ="Proj Fund Date"
    Expression ="SchGrp.FundDate_Projected"
    Alias ="Proj Synd Fee"
    Expression ="SchGrp.Fee_Projected"
    Alias ="Proj Group Term Penalty"
    Expression ="SchGrp.SgrpTermPenalty"
    Alias ="Proj Net Synd Penalty"
    Expression ="Nz([SchGrp].[Fee_Projected],0)+Nz([SchGrp].[sgrptermPenalty],0)"
    Alias ="Proj PD/U"
    Expression ="SchGrp.PerDiemProjected"
    Alias ="Proj UTA/U"
    Expression ="SchGrp.RebateUnitProjected"
    Alias ="Proj FMSS Fee/U"
    Expression ="SchGrp.SalesFSL"
    Alias ="Index -->"
    Expression ="\"\""
    Alias ="Proj Index"
    Expression ="SchGrp.SGrpSalesIndex"
    Alias ="Proj Date"
    Expression ="SchGrp.SGrpSalesIndDt"
    Alias ="Proj Term/Yr"
    Expression ="SchGrp.SGprSalesIndYr"
    Alias ="Proj Bonus"
    Expression ="SchGrp.SGrpSalesBonus"
    Alias ="Proj Rate"
    Expression ="SchGrp.SgrpSalesIndRt"
    Alias ="Proj Yield"
    Expression ="SchGrp.SgrpSalesYield"
    Alias ="UNIT -->"
    Expression ="\"\""
    Expression ="vw_SixKeys.AssetID"
    Alias ="Long Unit Description"
    Expression ="vw_SixKeys.PrintedDocs"
    Alias ="Unit Cost"
    Expression ="UnitTotalOEC.UnitCost"
    Alias ="Child Part Cost"
    Expression ="UnitTotalOEC.ChildPartCost"
    Alias ="Unit OEC"
    Expression ="UnitTotalOEC.UnitTotalCost"
    Alias ="Sch Rent"
    Expression ="Units.UnitSRent"
    Alias ="Sch RV"
    Expression ="Units.UnitSalesResid"
    Alias ="Synd Rent"
    Expression ="Units.UnitRent"
    Alias ="Synd RV"
    Expression ="Units.UnitSynResid"
    Alias ="LRF"
    Expression ="UnitTotalOEC.UnitLRF"
    Alias ="Not Needed? -->"
    Expression ="\"\""
    Expression ="SchGrp.LeasePaidAdvArr"
    Expression ="SchGrp.LeasePymtFreq"
    Alias ="HoldBack"
    Expression ="SchGrp.WithHeldPymts"
    Alias ="PDRent"
    Expression ="Units.UNITPDRENT"
    Expression ="vw_SixKeys.SchID"
    Expression ="vw_SixKeys.SGrpID"
End
Begin Joins
    LeftTable ="vw_SixKeys"
    RightTable ="UnitTotalOEC"
    Expression ="vw_SixKeys.AssetID = UnitTotalOEC.AssetID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="SchGrp"
    Expression ="vw_SixKeys.SGrpID = SchGrp.SGrpID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Schedule"
    Expression ="vw_SixKeys.SchID = Schedule.SchID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.SchID"
    Flag =0
    Expression ="vw_SixKeys.SGrpID"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Sch RV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.AssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Term"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Term/Yr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Synd Fee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MilesAllowed"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj PD/U"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLA# | Sch | Grp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Assignee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group Loan Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Index Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FMSS Fee/U"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sale Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group Synd Memo"
        dbInteger "ColumnWidth" ="3660"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group Service Fee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj UTA/U"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rate Used to Calc PV Service Fees"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SYND/FINAL ->"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Net Synd Penalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Economics Detail ->"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Index"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UTA/U"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sale Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Fund Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Broker Fee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Term/Yr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Misc 1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Group Term Penalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Fund Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Yield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Net Synd Penalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sch Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PD/U"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bonus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Downward"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Misc 2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Set By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Index ->"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Upward"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LO/TARGET -->"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Spread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Synd RV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rate Adjustment ->"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BP Change of"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Rate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Economics Detail -->"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Synd Fee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Group Term Penalty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj FMSS Fee/U"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Bonus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Index -->"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Index"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Child Part Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Proj Yield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UNIT -->"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Long Unit Description"
        dbInteger "ColumnWidth" ="8940"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit OEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Synd Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Not Needed? -->"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HoldBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.SchID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =18
    Top =-14
    Right =1299
    Bottom =972
    Left =-1
    Top =-1
    Right =1263
    Bottom =692
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="UnitTotalOEC"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="Units"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="Schedule"
        Name =""
    End
End
