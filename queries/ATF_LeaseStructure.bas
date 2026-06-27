Operation =1
Option =0
Begin InputTables
    Name ="ATF_Schedule_Group"
    Name ="ATF_UnitSpecCostRollup"
End
Begin OutputColumns
    Expression ="ATF_Schedule_Group.ATFID"
    Expression ="ATF_Schedule_Group.ATFSchName"
    Expression ="ATF_Schedule_Group.ATFGroupSpecID"
    Expression ="ATF_Schedule_Group.ATFGroupSpecDesc"
    Expression ="ATF_Schedule_Group.ATFUnitSpecID"
    Expression ="ATF_Schedule_Group.LeaseType"
    Expression ="ATF_Schedule_Group.PrimaryLeaseTerm"
    Expression ="ATF_Schedule_Group.PrimaryLeaseNotice"
    Expression ="ATF_Schedule_Group.LeasePymtFreq"
    Expression ="ATF_Schedule_Group.LeasePaidAdvArr"
    Expression ="ATF_Schedule_Group.ExchangeITGroup"
    Expression ="ATF_Schedule_Group.ExchangeITType"
    Expression ="ATF_Schedule_Group.ExchangeITMonth"
    Expression ="ATF_Schedule_Group.ExtendITOptionTerm"
    Expression ="ATF_Schedule_Group.ExtendedRentAvg"
    Expression ="ATF_Schedule_Group.MTMRentAvg"
    Expression ="ATF_Schedule_Group.SwapIndex"
    Expression ="ATF_Schedule_Group.SwapDate"
    Expression ="ATF_Schedule_Group.SwapRate"
    Expression ="ATF_Schedule_Group.SwapTermYears"
    Expression ="ATF_Schedule_Group.SwapSpread"
    Expression ="ATF_Schedule_Group.SwapYield"
    Expression ="ATF_Schedule_Group.SwapBonus"
    Expression ="ATF_Schedule_Group.SwapAdjustment"
    Expression ="ATF_Schedule_Group.ResidualAvg"
    Alias ="NegBLCD"
    Expression ="DLookUp(\"StartsOn\",\"Lookup_LeaseStartsOn\",\"ID=\" & [BLCDBasisID])"
    Alias ="PerDiemType"
    Expression ="DLookUp(\"StartsOn\",\"Lookup_LeaseStartsOn\",\"ID=\" & [PerDiemBasisID])"
    Expression ="ATF_UnitSpecCostRollup.UnitCost"
    Expression ="ATF_UnitSpecCostRollup.PartCost"
    Expression ="ATF_UnitSpecCostRollup.TotalUnitCost"
    Alias ="UAT"
    Expression ="CCur(Nz([UATPerUnit],0))"
    Expression ="ATF_UnitSpecCostRollup.PerDiemRent"
    Expression ="ATF_Schedule_Group.SchRent"
    Expression ="ATF_UnitSpecCostRollup.LRFSch"
    Expression ="ATF_Schedule_Group.FirstExtendedRent"
    Expression ="ATF_UnitSpecCostRollup.LRFExt"
    Expression ="ATF_Schedule_Group.MTMRent"
    Expression ="ATF_UnitSpecCostRollup.LRFMTM"
    Alias ="Units"
    Expression ="ATF_UnitSpecCostRollup.NumUnitsInSpec"
End
Begin Joins
    LeftTable ="ATF_Schedule_Group"
    RightTable ="ATF_UnitSpecCostRollup"
    Expression ="ATF_Schedule_Group.ATFGroupSpecID = ATF_UnitSpecCostRollup.ATFGroupSpecID"
    Flag =2
End
Begin Groups
    Expression ="ATF_Schedule_Group.ATFID"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.ATFSchName"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.ATFGroupSpecID"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.ATFGroupSpecDesc"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.ATFUnitSpecID"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.LeaseType"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.PrimaryLeaseTerm"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.PrimaryLeaseNotice"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.LeasePymtFreq"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.LeasePaidAdvArr"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.ExchangeITGroup"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.ExchangeITType"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.ExchangeITMonth"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.ExtendITOptionTerm"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.ExtendedRentAvg"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.MTMRentAvg"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.SwapIndex"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.SwapDate"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.SwapRate"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.SwapTermYears"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.SwapSpread"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.SwapYield"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.SwapBonus"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.SwapAdjustment"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.ResidualAvg"
    GroupLevel =0
    Expression ="DLookUp(\"StartsOn\",\"Lookup_LeaseStartsOn\",\"ID=\" & [BLCDBasisID])"
    GroupLevel =0
    Expression ="DLookUp(\"StartsOn\",\"Lookup_LeaseStartsOn\",\"ID=\" & [PerDiemBasisID])"
    GroupLevel =0
    Expression ="ATF_UnitSpecCostRollup.UnitCost"
    GroupLevel =0
    Expression ="ATF_UnitSpecCostRollup.PartCost"
    GroupLevel =0
    Expression ="ATF_UnitSpecCostRollup.TotalUnitCost"
    GroupLevel =0
    Expression ="CCur(Nz([UATPerUnit],0))"
    GroupLevel =0
    Expression ="ATF_UnitSpecCostRollup.PerDiemRent"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.SchRent"
    GroupLevel =0
    Expression ="ATF_UnitSpecCostRollup.LRFSch"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.FirstExtendedRent"
    GroupLevel =0
    Expression ="ATF_UnitSpecCostRollup.LRFExt"
    GroupLevel =0
    Expression ="ATF_Schedule_Group.MTMRent"
    GroupLevel =0
    Expression ="ATF_UnitSpecCostRollup.LRFMTM"
    GroupLevel =0
    Expression ="ATF_UnitSpecCostRollup.NumUnitsInSpec"
    GroupLevel =0
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
        dbText "Name" ="ATF_Schedule_Group.SwapAdjustment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.PrimaryLeaseNotice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.ResidualAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.ExchangeITGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.PrimaryLeaseTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.ExchangeITType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.ExchangeITMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.ExtendITOptionTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.ExtendedRentAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.MTMRentAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.SwapIndex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.SwapDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.SwapRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.SwapTermYears"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.SwapSpread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.SwapYield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NegBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiemType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.ATFUnitSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UAT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.SchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.MTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.FirstExtendedRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.ATFGroupSpecDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.LeaseType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_UnitSpecCostRollup.LRFSch"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_UnitSpecCostRollup.PartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_UnitSpecCostRollup.TotalUnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_UnitSpecCostRollup.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_UnitSpecCostRollup.LRFExt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_UnitSpecCostRollup.LRFMTM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_UnitSpecCostRollup.PerDiemRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.ATFSchName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATF_Schedule_Group.SwapBonus"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =60
    Top =19
    Right =1352
    Bottom =790
    Left =-1
    Top =-1
    Right =1274
    Bottom =90
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =47
        Top =0
        Right =292
        Bottom =454
        Top =0
        Name ="ATF_Schedule_Group"
        Name =""
    End
    Begin
        Left =492
        Top =160
        Right =636
        Bottom =304
        Top =0
        Name ="ATF_UnitSpecCostRollup"
        Name =""
    End
End
