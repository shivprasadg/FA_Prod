Operation =1
Option =0
Where ="(((ScheduleUnitDocs3.AssetId) Is Not Null))"
Begin InputTables
    Name ="ScheduleUnitDocs3"
    Name ="Cost_FeesTax"
    Name ="Cost_Freight"
    Name ="Cost_Refrig"
    Name ="Cost_Warranty"
    Name ="Cost_Other"
    Name ="Cost_PTO"
    Name ="Cost_Inspection"
    Name ="Cost_DecalsCoName"
    Name ="Cost_DecalsRegNumber"
    Name ="Cost_TankBody"
    Name ="Cost_TankOther"
    Name ="Cost_Oil"
    Name ="Cost_DecalsMisc"
    Name ="Cost_OBC_All"
    Name ="Cost_DecalsMisc2"
End
Begin OutputColumns
    Expression ="ScheduleUnitDocs3.[Lease No]"
    Expression ="ScheduleUnitDocs3.Lessee"
    Expression ="ScheduleUnitDocs3.Co_Lessee"
    Expression ="ScheduleUnitDocs3.SchNo"
    Expression ="ScheduleUnitDocs3.Group"
    Expression ="ScheduleUnitDocs3.[Unit No]"
    Expression ="ScheduleUnitDocs3.Unitvin"
    Expression ="ScheduleUnitDocs3.Desc"
    Expression ="ScheduleUnitDocs3.UnitCost"
    Expression ="ScheduleUnitDocs3.UnitInvoiceNum"
    Expression ="ScheduleUnitDocs3.UnitInvoicePaid"
    Alias ="RefrigMake"
    Expression ="Cost_Refrig.RefrigMake"
    Alias ="RefrigSerial"
    Expression ="Cost_Refrig.RefrigSerialNum"
    Alias ="Refrig"
    Expression ="CCur(Nz([Cost_Refrig].[Refrig],0))"
    Expression ="Cost_Refrig.RefrigInvoice"
    Expression ="Cost_Refrig.RefrigPaid"
    Expression ="Cost_FeesTax.TaxName"
    Alias ="FeesTax"
    Expression ="CCur(Nz([Cost_FeesTax].[FeesTax],0))"
    Expression ="Cost_FeesTax.TaxInvoice"
    Expression ="Cost_FeesTax.TaxPaid"
    Expression ="Cost_Freight.FreightName"
    Alias ="Freight"
    Expression ="CCur(Nz([Cost_Freight].[Freight],0))"
    Expression ="Cost_Freight.FGTInvoice"
    Expression ="Cost_Freight.FreightPaid"
    Expression ="Cost_Warranty.WarrantyName"
    Alias ="Warranty"
    Expression ="CCur(Nz([Cost_Warranty].[Warranty],0))"
    Expression ="Cost_Warranty.WarrantyInv"
    Expression ="Cost_Warranty.WarrantyPaid"
    Expression ="Cost_Inspection.InspCo"
    Alias ="Inspection"
    Expression ="CCur(Nz([Cost_Inspection].[Inspection],0))"
    Expression ="Cost_Inspection.InspInvoice"
    Expression ="Cost_Inspection.InspPaid"
    Expression ="Cost_Other.OtherMake"
    Alias ="Other"
    Expression ="CCur(Nz([Cost_Other].[other],0))"
    Expression ="Cost_Other.OtherInvoice"
    Expression ="Cost_Other.OtherPaid"
    Expression ="Cost_PTO.PTOMake"
    Alias ="PTO"
    Expression ="CCur(Nz([Cost_PTO].[PTO],0))"
    Expression ="Cost_PTO.PTOInvoice"
    Expression ="Cost_PTO.PTOPaid"
    Alias ="CoNameDecal"
    Expression ="Cost_DecalsCoName.Name"
    Alias ="CoNameCost"
    Expression ="CCur(Nz([Cost_DecalsCoName].[CoNameCost],0))"
    Expression ="Cost_DecalsCoName.CoNameInvoice"
    Expression ="Cost_DecalsCoName.CoNamePaid"
    Alias ="RegDecal"
    Expression ="Cost_DecalsRegNumber.Name"
    Alias ="RegDecalCost"
    Expression ="CCur(Nz([Cost_DecalsRegNumber].[RegDecalCost],0))"
    Expression ="Cost_DecalsRegNumber.RegDecalInvoice"
    Expression ="Cost_DecalsRegNumber.RegDecalPaid"
    Alias ="MiscDecal"
    Expression ="Cost_DecalsMisc.Name"
    Alias ="MiscCost"
    Expression ="CCur(Nz([Cost_DecalsMisc].[MiscCost],0))"
    Expression ="Cost_DecalsMisc.MiscInvoice"
    Expression ="Cost_DecalsMisc.MiscPaid"
    Alias ="Misc2Decal"
    Expression ="Cost_DecalsMisc2.Name"
    Alias ="Misc2Cost"
    Expression ="CCur(Nz([Cost_DecalsMisc2].[MiscCost],0))"
    Alias ="Misc2Invoice"
    Expression ="Cost_DecalsMisc2.MiscInvoice"
    Alias ="Misc2Paid"
    Expression ="Cost_DecalsMisc2.MiscPaid"
    Expression ="Cost_TankBody.TankMake"
    Alias ="TankBody"
    Expression ="CCur(Nz([Cost_TankBody].[tankbody],0))"
    Expression ="Cost_TankBody.TankInvoice"
    Expression ="Cost_TankBody.TankPaid"
    Expression ="Cost_TankOther.TankOthMake"
    Alias ="TankOther"
    Expression ="CCur(Nz([Cost_TankOther].[tankOther],0))"
    Expression ="Cost_TankOther.TankOthInvoice"
    Expression ="Cost_TankOther.TankOthPaid"
    Expression ="Cost_Oil.OilRigMake"
    Alias ="OilRig"
    Expression ="CCur(Nz([Cost_Oil].[OilRig],0))"
    Expression ="Cost_Oil.OilRigInvoice"
    Expression ="Cost_Oil.OilRigPaid"
    Expression ="Cost_OBC_All.OBC_Make"
    Expression ="Cost_OBC_All.OBC_Cost"
    Expression ="Cost_OBC_All.OBCInvoice"
    Expression ="Cost_OBC_All.OBCPaid"
    Expression ="Cost_OBC_All.OBC_Display"
    Expression ="Cost_OBC_All.OBCDisplay"
    Expression ="Cost_OBC_All.OBCDisplayInv"
    Expression ="Cost_OBC_All.OBCDisplayPaid"
    Expression ="Cost_OBC_All.OBC_Adapt"
    Expression ="Cost_OBC_All.OBCAdaptr"
    Expression ="Cost_OBC_All.OBCAdaptInv"
    Expression ="Cost_OBC_All.OBCAdaptPaid"
    Expression ="Cost_OBC_All.OBC_Dome"
    Expression ="Cost_OBC_All.OBCDome"
    Expression ="Cost_OBC_All.OBCDomeInv"
    Expression ="Cost_OBC_All.OBCDomePaid"
    Expression ="Cost_OBC_All.OBC_Multi"
    Expression ="Cost_OBC_All.OBCMulti"
    Expression ="Cost_OBC_All.OBCMultiInv"
    Expression ="Cost_OBC_All.OBCMultiPaid"
    Expression ="Cost_OBC_All.OBCInstaller"
    Expression ="Cost_OBC_All.OBCInstall"
    Expression ="Cost_OBC_All.OBCInsInvoice"
    Expression ="Cost_OBC_All.OBCInstallPaid"
    Alias ="OBCPartMake"
    Expression ="Cost_OBC_All.OBC_Part"
    Expression ="Cost_OBC_All.OBCPart"
    Expression ="Cost_OBC_All.OBCPartInv"
    Expression ="Cost_OBC_All.OBCPartPaid"
    Alias ="UnitTotalCost"
    Expression ="CCur([UnitCost]+[Refrig]+[OBCTotal]+[FeesTax]+[Freight]+[Warranty]+[Other]+[Insp"
        "ection]+[PTO]+[CoNameCost]+[RegDecalCost]+[MiscCost]+[Misc2Cost]+[TankBody]+[Tan"
        "kOther]+[OilRig])"
    Expression ="ScheduleUnitDocs3.AssetId"
    Expression ="ScheduleUnitDocs3.Status"
    Expression ="ScheduleUnitDocs3.InService"
    Expression ="ScheduleUnitDocs3.ActualDelDate"
    Expression ="ScheduleUnitDocs3.[Title Num]"
    Expression ="ScheduleUnitDocs3.TitledState"
    Expression ="ScheduleUnitDocs3.Location"
    Expression ="ScheduleUnitDocs3.StreetAddress"
    Expression ="ScheduleUnitDocs3.City"
    Expression ="ScheduleUnitDocs3.State"
    Expression ="ScheduleUnitDocs3.Zip"
    Expression ="ScheduleUnitDocs3.InterimBank"
    Expression ="ScheduleUnitDocs3.InterimDate"
    Expression ="ScheduleUnitDocs3.InterimAmt"
End
Begin Joins
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_FeesTax"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_FeesTax.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_Freight"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_Freight.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_Refrig"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_Refrig.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_Warranty"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_Warranty.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_Other"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_Other.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_PTO"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_PTO.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_Inspection"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_Inspection.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_DecalsCoName"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_DecalsCoName.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_DecalsRegNumber"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_DecalsRegNumber.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_TankBody"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_TankBody.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_TankOther"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_TankOther.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_Oil"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_Oil.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_DecalsMisc"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_DecalsMisc.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_OBC_All"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_OBC_All.AssetId"
    Flag =2
    LeftTable ="ScheduleUnitDocs3"
    RightTable ="Cost_DecalsMisc2"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_DecalsMisc2.AssetId"
    Flag =2
End
Begin OrderBy
    Expression ="ScheduleUnitDocs3.Group"
    Flag =0
    Expression ="ScheduleUnitDocs3.[Unit No]"
    Flag =0
    Expression ="ScheduleUnitDocs3.AssetId"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="FeesTax"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Freight"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Refrig"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Lease No"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Group"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.AssetId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Unit No"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Lessee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Co_Lessee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Location"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="68"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.StreetAddress"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="69"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.City"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="70"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.State"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="71"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Zip"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="72"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Title Num"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="61"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="65"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.InService"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="66"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.UnitCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Other"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="Cost_Other.OtherMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="Cost_Other.OtherInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="Cost_Inspection.InspCo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="Cost_Inspection.InspInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="Inspection"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="Cost_PTO.PTOMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="Cost_PTO.PTOInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="38"
    End
    Begin
        dbText "Name" ="PTO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="Cost_Inspection.InspPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="Cost_PTO.PTOPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="39"
    End
    Begin
        dbText "Name" ="Cost_DecalsCoName.CoNameInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="42"
    End
    Begin
        dbText "Name" ="Cost_DecalsCoName.CoNamePaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="43"
    End
    Begin
        dbText "Name" ="Cost_DecalsRegNumber.RegDecalInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="46"
    End
    Begin
        dbText "Name" ="Cost_DecalsRegNumber.RegDecalPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="47"
    End
    Begin
        dbText "Name" ="CoNameDecal"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="40"
    End
    Begin
        dbText "Name" ="RegDecal"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="44"
    End
    Begin
        dbText "Name" ="Cost_DecalsMisc.MiscInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="50"
    End
    Begin
        dbText "Name" ="Cost_DecalsMisc.MiscPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="51"
    End
    Begin
        dbText "Name" ="MiscDecal"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="48"
    End
    Begin
        dbText "Name" ="CoNameCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="41"
    End
    Begin
        dbText "Name" ="RegDecalCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="45"
    End
    Begin
        dbText "Name" ="MiscCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="49"
    End
    Begin
        dbText "Name" ="Cost_FeesTax.TaxName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="Cost_FeesTax.TaxInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="30"
    End
    Begin
        dbText "Name" ="Cost_FeesTax.TaxPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="Cost_Freight.FreightName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="Cost_Freight.FGTInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="33"
    End
    Begin
        dbText "Name" ="Cost_Freight.FreightPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="34"
    End
    Begin
        dbText "Name" ="Cost_Warranty.WarrantyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="35"
    End
    Begin
        dbText "Name" ="Cost_Warranty.WarrantyInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="36"
    End
    Begin
        dbText "Name" ="Cost_Warranty.WarrantyPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="37"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.[Lease No]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.[Unit No]"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Unitvin"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.[Title Num]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="67"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.UnitInvoicePaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Cost_Other.OtherPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="Cost_TankBody.TankMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="52"
    End
    Begin
        dbText "Name" ="Cost_TankBody.TankInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="54"
    End
    Begin
        dbText "Name" ="Cost_TankBody.TankPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="55"
    End
    Begin
        dbText "Name" ="Cost_TankOther.TankOthPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="59"
    End
    Begin
        dbText "Name" ="Cost_Oil.OilRigPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="63"
    End
    Begin
        dbText "Name" ="Cost_Oil.OilRigMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="60"
    End
    Begin
        dbText "Name" ="Cost_TankOther.TankOthMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="56"
    End
    Begin
        dbText "Name" ="Cost_TankOther.TankOthInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="58"
    End
    Begin
        dbText "Name" ="Cost_Oil.OilRigInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="62"
    End
    Begin
        dbText "Name" ="TankBody"
        dbInteger "ColumnOrder" ="53"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankOther"
        dbInteger "ColumnOrder" ="57"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OilRig"
        dbInteger "ColumnOrder" ="61"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Make"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="75"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Cost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="76"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="77"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="78"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Adapt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="79"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCAdaptr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="80"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCAdaptInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="81"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCAdaptPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="82"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Dome"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="83"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCDome"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="84"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCDomeInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="85"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCDomePaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="86"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Multi"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="87"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCMulti"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="88"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCMultiInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="89"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCMultiPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="90"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInstaller"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="91"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInstall"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="92"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInsInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="93"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInstallPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="94"
    End
    Begin
        dbText "Name" ="RefrigMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="73"
    End
    Begin
        dbText "Name" ="RefrigSerial"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="74"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.InterimBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.InterimDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.InterimAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.titledstate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.ActualDelDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="95"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Display"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCDisplay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCDisplayInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCDisplayPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Misc2Cost"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Misc2Decal"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Misc2Paid"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Misc2Invoice"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCPartPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCPart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCPartInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPartMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalCost"
        dbInteger "ColumnOrder" ="64"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TaxPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FreightName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FGTInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FreightPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WarrantyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WarrantyInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="WarrantyPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InspCo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InspInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InspPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OtherMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OtherInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OtherPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PTOMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PTOInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PTOPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CoNameInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CoNamePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegDecalInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RegDecalPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MiscInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MiscPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankOthMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankOthInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TankOthPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OilRigMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OilRigInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OilRigPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC_Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC_Cost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC_Display"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDisplay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDisplayInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDisplayPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC_Adapt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCAdaptr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCAdaptInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCAdaptPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC_Dome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDomeInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCDomePaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBC_Multi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCMulti"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCMultiInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCMultiPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInstaller"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInstall"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInsInvoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCInstallPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPartInv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OBCPartPaid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title Num"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =201
    Top =26
    Right =2147
    Bottom =838
    Left =-1
    Top =-1
    Right =1914
    Bottom =252
    Left =475
    Top =0
    ColumnsShown =539
    Begin
        Left =-447
        Top =8
        Right =-281
        Bottom =352
        Top =0
        Name ="ScheduleUnitDocs3"
        Name =""
    End
    Begin
        Left =-258
        Top =17
        Right =-160
        Bottom =161
        Top =0
        Name ="Cost_FeesTax"
        Name =""
    End
    Begin
        Left =-147
        Top =19
        Right =-47
        Bottom =163
        Top =0
        Name ="Cost_Freight"
        Name =""
    End
    Begin
        Left =227
        Top =25
        Right =353
        Bottom =169
        Top =0
        Name ="Cost_Refrig"
        Name =""
    End
    Begin
        Left =369
        Top =26
        Right =475
        Bottom =174
        Top =0
        Name ="Cost_Warranty"
        Name =""
    End
    Begin
        Left =776
        Top =37
        Right =926
        Bottom =269
        Top =0
        Name ="Cost_Other"
        Name =""
    End
    Begin
        Left =483
        Top =28
        Right =591
        Bottom =172
        Top =0
        Name ="Cost_PTO"
        Name =""
    End
    Begin
        Left =603
        Top =32
        Right =714
        Bottom =176
        Top =0
        Name ="Cost_Inspection"
        Name =""
    End
    Begin
        Left =-254
        Top =178
        Right =-111
        Bottom =338
        Top =0
        Name ="Cost_DecalsCoName"
        Name =""
    End
    Begin
        Left =-102
        Top =181
        Right =51
        Bottom =341
        Top =0
        Name ="Cost_DecalsRegNumber"
        Name =""
    End
    Begin
        Left =206
        Top =184
        Right =329
        Bottom =326
        Top =0
        Name ="Cost_TankBody"
        Name =""
    End
    Begin
        Left =347
        Top =183
        Right =466
        Bottom =329
        Top =0
        Name ="Cost_TankOther"
        Name =""
    End
    Begin
        Left =484
        Top =183
        Right =594
        Bottom =328
        Top =0
        Name ="Cost_Oil"
        Name =""
    End
    Begin
        Left =72
        Top =181
        Right =193
        Bottom =344
        Top =0
        Name ="Cost_DecalsMisc"
        Name =""
    End
    Begin
        Left =1
        Top =12
        Right =145
        Bottom =156
        Top =0
        Name ="Cost_OBC_All"
        Name =""
    End
    Begin
        Left =604
        Top =185
        Right =753
        Bottom =334
        Top =0
        Name ="Cost_DecalsMisc2"
        Name =""
    End
End
