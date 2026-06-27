Operation =1
Option =0
Begin InputTables
    Name ="ATFScheduleSpec"
    Name ="ATFGroupSpec"
    Name ="ATFUnitSpec"
    Name ="ATFChildPartSpec"
End
Begin OutputColumns
    Expression ="ATFScheduleSpec.ATFID"
    Expression ="ATFScheduleSpec.MLID"
    Expression ="ATFScheduleSpec.ATFSchSpecID"
    Expression ="ATFGroupSpec.ATFGroupSpecID"
    Expression ="ATFScheduleSpec.ATFSchName"
    Expression ="ATFScheduleSpec.ATFSchSpecDesc"
    Expression ="ATFScheduleSpec.SchFAID"
    Expression ="ATFScheduleSpec.FSLType"
    Expression ="ATFScheduleSpec.SchDescription"
    Expression ="ATFScheduleSpec.NumGroupsInSch"
    Expression ="ATFScheduleSpec.ApprovalRequest"
    Expression ="ATFScheduleSpec.ApprovalReply"
    Expression ="ATFScheduleSpec.ScheduleBillingID"
    Expression ="ATFScheduleSpec.ATFSchNotes"
    Expression ="ATFScheduleSpec.SchSpecInserted"
    Expression ="ATFGroupSpec.ATFGroupSpecDesc"
    Expression ="ATFGroupSpec.UnitSpecsInGroup"
    Expression ="ATFGroupSpec.UnitsInGroup"
    Expression ="ATFGroupSpec.GroupOEC"
    Expression ="ATFGroupSpec.GrpFAID"
    Expression ="ATFGroupSpec.ExchangeITGroup"
    Expression ="ATFGroupSpec.ExchangeITType"
    Expression ="ATFGroupSpec.ExchangeITMonth"
    Expression ="ATFGroupSpec.ExtendITOptionTerm"
    Expression ="ATFGroupSpec.ExtendedRentAvg"
    Expression ="ATFGroupSpec.MTMRentAvg"
    Expression ="ATFGroupSpec.PrimaryLeaseTerm"
    Expression ="ATFGroupSpec.PrimaryLeaseNotice"
    Expression ="ATFGroupSpec.LeasePymtFreq"
    Expression ="ATFGroupSpec.LeasePaidAdvArr"
    Expression ="ATFGroupSpec.SchRentAvg"
    Expression ="ATFGroupSpec.LRFSchGroup"
    Expression ="ATFGroupSpec.SyndRentAvg"
    Expression ="ATFGroupSpec.LRFSyndGroup"
    Expression ="ATFGroupSpec.SwapIndex"
    Expression ="ATFGroupSpec.SwapDate"
    Expression ="ATFGroupSpec.SwapRate"
    Expression ="ATFGroupSpec.SwapTermYears"
    Expression ="ATFGroupSpec.SwapSpread"
    Expression ="ATFGroupSpec.SwapYield"
    Expression ="ATFGroupSpec.SwapBonus"
    Expression ="ATFGroupSpec.SwapAdjustment"
    Expression ="ATFGroupSpec.ResidualAvg"
    Expression ="ATFGroupSpec.ResidualPctAvg"
    Expression ="ATFGroupSpec.UATPerUnit"
    Expression ="ATFGroupSpec.PLB"
    Expression ="ATFGroupSpec.LeaseType"
    Expression ="ATFGroupSpec.BLCDBasisID"
    Expression ="ATFGroupSpec.PerDiemBasisID"
    Expression ="ATFGroupSpec.AllowanceMethod"
    Expression ="ATFGroupSpec.MPYAllowance"
    Expression ="ATFGroupSpec.CPMile"
    Expression ="ATFGroupSpec.CPMileThreshold"
    Expression ="ATFGroupSpec.CPMileAfter"
    Expression ="ATFGroupSpec.HPYAllowance"
    Expression ="ATFGroupSpec.CPHour"
    Expression ="ATFGroupSpec.CPHourThreshold"
    Expression ="ATFGroupSpec.CPHourAfter"
    Expression ="ATFGroupSpec.FMSFeeAmount"
    Expression ="ATFGroupSpec.FMSFeeStartMonth"
    Expression ="ATFGroupSpec.FMSFeeTermMonths"
    Expression ="ATFGroupSpec.Fee2Type"
    Expression ="ATFGroupSpec.Fee2Amount"
    Expression ="ATFGroupSpec.Fee2TermMonths"
    Expression ="ATFGroupSpec.Fee2StartMonth"
    Expression ="ATFGroupSpec.ATFGroupNotes"
    Expression ="ATFGroupSpec.GroupSpecInserted"
    Expression ="ATFGroupSpec.LastModified"
    Expression ="ATFUnitSpec.ATFUnitSpecDesc"
    Expression ="ATFUnitSpec.ATFUnitSpecID"
    Expression ="ATFUnitSpec.NumUnitsInSpec"
    Expression ="ATFUnitSpec.UnitModelYear"
    Expression ="ATFUnitSpec.UnitMake"
    Expression ="ATFUnitSpec.UnitModel"
    Expression ="ATFUnitSpec.UnitType"
    Expression ="ATFUnitSpec.UnitSubType"
    Expression ="ATFUnitSpec.UOM"
    Expression ="ATFUnitSpec.UnitConfiguration"
    Expression ="ATFUnitSpec.Axle"
    Expression ="ATFUnitSpec.AxleProperty"
    Expression ="ATFUnitSpec.PowerType"
    Expression ="ATFUnitSpec.UnitDescriptor"
    Expression ="ATFUnitSpec.EngineType"
    Expression ="ATFUnitSpec.TransmissionType"
    Expression ="ATFUnitSpec.GVW"
    Expression ="ATFUnitSpec.VendorID"
    Expression ="ATFUnitSpec.UnitCost"
    Expression ="ATFUnitSpec.PerDiemRent"
    Expression ="ATFUnitSpec.SchRent"
    Expression ="ATFUnitSpec.SyndRent"
    Expression ="ATFUnitSpec.FirstExtendedRent"
    Expression ="ATFUnitSpec.FirstExtendedTerm"
    Expression ="ATFUnitSpec.MTMRent"
    Expression ="ATFUnitSpec.SalesResidual"
    Expression ="ATFUnitSpec.SyndResidual"
    Expression ="ATFUnitSpec.EstimatedBuildDate"
    Expression ="ATFUnitSpec.EstimatedDeliveryDate"
    Expression ="ATFUnitSpec.DeliveryLocationID"
    Expression ="ATFUnitSpec.GarageLocationID"
    Expression ="ATFUnitSpec.RegTitleLocationID"
    Expression ="ATFUnitSpec.BillingContactID"
    Expression ="ATFUnitSpec.SpecInserted"
    Expression ="ATFUnitSpec.GroupID"
    Expression ="ATFUnitSpec.ClientGroupID"
    Expression ="ATFUnitSpec.ClientID"
    Expression ="ATFUnitSpec.LegalDescription"
    Expression ="ATFUnitSpec.NumOfChildSpecs"
    Expression ="ATFUnitSpec.LastModified"
    Expression ="ATFUnitSpec.UnitSpecNotes"
    Expression ="ATFUnitSpec.SpecProvidedBy"
    Expression ="ATFChildPartSpec.ATFChildSpecID"
    Expression ="ATFChildPartSpec.ATFChildSpecDesc"
    Expression ="ATFChildPartSpec.ChildVendorID"
    Expression ="ATFChildPartSpec.ChildPartYear"
    Expression ="ATFChildPartSpec.ChildMake"
    Expression ="ATFChildPartSpec.ChildClass"
    Expression ="ATFChildPartSpec.ChildType"
    Expression ="ATFChildPartSpec.ChildPartCost"
    Expression ="ATFChildPartSpec.ChildpartSpecNotes"
    Expression ="ATFChildPartSpec.ChildSpecInserted"
    Expression ="ATFChildPartSpec.LastModified"
    Expression ="ATFUnitSpec.UTARebate"
    Expression ="ATFGroupSpec.Fee_Projected"
End
Begin Joins
    LeftTable ="ATFScheduleSpec"
    RightTable ="ATFGroupSpec"
    Expression ="ATFScheduleSpec.ATFSchSpecID = ATFGroupSpec.ATFSchSpecID"
    Flag =2
    LeftTable ="ATFGroupSpec"
    RightTable ="ATFUnitSpec"
    Expression ="ATFGroupSpec.ATFGroupSpecID = ATFUnitSpec.ATFGroupSpecID"
    Flag =2
    LeftTable ="ATFUnitSpec"
    RightTable ="ATFChildPartSpec"
    Expression ="ATFUnitSpec.ATFUnitSpecID = ATFChildPartSpec.ATFUnitSpecID"
    Flag =2
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
        dbText "Name" ="ATFGroupSpec.Fee2Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.SchSpecInserted"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.BLCDBasisID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.LastModified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.PrimaryLeaseNotice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.ATFSchSpecDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.ATFGroupSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.SchFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.FMSFeeTermMonths"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.ATFID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.LRFSyndGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.ExtendITOptionTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.CPMileThreshold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.ResidualAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.ATFSchNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.ATFSchSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.HPYAllowance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.UATPerUnit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.ATFSchName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.SwapAdjustment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.GrpFAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.FSLType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.ResidualPctAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.SchDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.ExchangeITType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.NumGroupsInSch"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.SwapIndex"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.ApprovalRequest"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.ApprovalReply"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFScheduleSpec.ScheduleBillingID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.ATFGroupSpecDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.UnitSpecsInGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.UnitsInGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.LeasePymtFreq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.CPHour"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.PLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.GroupOEC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.ExchangeITGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.ExchangeITMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.SwapRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.ExtendedRentAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.SwapSpread"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.MTMRentAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.PrimaryLeaseTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.LeasePaidAdvArr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.Fee2TermMonths"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.SchRentAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.SwapYield"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.LRFSchGroup"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.SyndRentAvg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.SwapDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.SwapTermYears"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.LeaseType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.PerDiemBasisID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.AllowanceMethod"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.MPYAllowance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.CPMile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.CPMileAfter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.CPHourThreshold"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.CPHourAfter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.FMSFeeAmount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.FMSFeeStartMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.Fee2Amount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.Fee2StartMonth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.ATFGroupNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.GroupSpecInserted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.LastModified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.NumUnitsInSpec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.PowerType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SpecProvidedBy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ATFUnitSpecDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ATFUnitSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitModelYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UOM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitConfiguration"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.AxleProperty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitDescriptor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.EngineType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.TransmissionType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.GVW"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SpecInserted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.PerDiemRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SchRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.EstimatedBuildDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.FirstExtendedRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.FirstExtendedTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.MTMRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SalesResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.SyndResidual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.EstimatedDeliveryDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.DeliveryLocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.GarageLocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.RegTitleLocationID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.BillingContactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ClientGroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.LegalDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.NumOfChildSpecs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UnitSpecNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.LastModified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFUnitSpec.UTARebate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.SwapBonus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ATFChildSpecID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ATFChildSpecDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ChildVendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ChildPartYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ChildMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ChildModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ChildpartSpecNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFChildPartSpec.ChildSpecInserted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ATFGroupSpec.Fee_Projected"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =178
    Top =22
    Right =1702
    Bottom =1008
    Left =-1
    Top =-1
    Right =1500
    Bottom =607
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =122
        Top =71
        Right =266
        Bottom =415
        Top =0
        Name ="ATFScheduleSpec"
        Name =""
    End
    Begin
        Left =310
        Top =177
        Right =454
        Bottom =666
        Top =0
        Name ="ATFGroupSpec"
        Name =""
    End
    Begin
        Left =535
        Top =188
        Right =679
        Bottom =636
        Top =0
        Name ="ATFUnitSpec"
        Name =""
    End
    Begin
        Left =775
        Top =228
        Right =919
        Bottom =372
        Top =0
        Name ="ATFChildPartSpec"
        Name =""
    End
End
