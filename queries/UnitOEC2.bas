Operation =1
Option =0
Where ="(((UnitOEC1.AssetId) Is Not Null))"
Begin InputTables
    Name ="UnitOEC1"
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
End
Begin OutputColumns
    Expression ="UnitOEC1.AssetId"
    Expression ="UnitOEC1.UnitNo"
    Expression ="UnitOEC1.Unitvin"
    Expression ="UnitOEC1.Status"
    Expression ="UnitOEC1.GroupREF"
    Expression ="UnitOEC1.Desc"
    Expression ="UnitOEC1.UnitCost"
    Alias ="TotalCost"
    Expression ="CCur([UnitCost]+[Refrig]+[OBCTotal]+[FeesTax]+[Freight]+[Warranty]+[Other]+[Insp"
        "ection]+[PTO]+[CoNameCost]+[RegDecalCost]+[MiscCost]+[TankBody]+[TankOther]+[Oil"
        "Rig])"
    Expression ="UnitOEC1.UnitInvoiceNum"
    Expression ="UnitOEC1.UnitInvoicePaid"
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
    Expression ="UnitOEC1.InService"
    Expression ="UnitOEC1.[Title Num]"
    Expression ="UnitOEC1.TitledState"
    Expression ="UnitOEC1.Location"
    Expression ="UnitOEC1.StreetAddress"
    Expression ="UnitOEC1.City"
    Expression ="UnitOEC1.State"
    Expression ="UnitOEC1.Zip"
    Expression ="UnitOEC1.InterimBank"
    Expression ="UnitOEC1.InterimDate"
    Expression ="UnitOEC1.InterimAmt"
    Expression ="UnitOEC1.DrawDownId"
    Expression ="UnitOEC1.VendorREF"
    Expression ="UnitOEC1.PerDiemRent"
    Expression ="UnitOEC1.LORent"
    Expression ="UnitOEC1.SYNDRent"
End
Begin Joins
    LeftTable ="UnitOEC1"
    RightTable ="Cost_FeesTax"
    Expression ="UnitOEC1.AssetId = Cost_FeesTax.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_Freight"
    Expression ="UnitOEC1.AssetId = Cost_Freight.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_Refrig"
    Expression ="UnitOEC1.AssetId = Cost_Refrig.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_Warranty"
    Expression ="UnitOEC1.AssetId = Cost_Warranty.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_Other"
    Expression ="UnitOEC1.AssetId = Cost_Other.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_PTO"
    Expression ="UnitOEC1.AssetId = Cost_PTO.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_Inspection"
    Expression ="UnitOEC1.AssetId = Cost_Inspection.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_DecalsCoName"
    Expression ="UnitOEC1.AssetId = Cost_DecalsCoName.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_DecalsRegNumber"
    Expression ="UnitOEC1.AssetId = Cost_DecalsRegNumber.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_TankBody"
    Expression ="UnitOEC1.AssetId = Cost_TankBody.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_TankOther"
    Expression ="UnitOEC1.AssetId = Cost_TankOther.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_Oil"
    Expression ="UnitOEC1.AssetId = Cost_Oil.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_DecalsMisc"
    Expression ="UnitOEC1.AssetId = Cost_DecalsMisc.AssetId"
    Flag =2
    LeftTable ="UnitOEC1"
    RightTable ="Cost_OBC_All"
    Expression ="UnitOEC1.AssetId = Cost_OBC_All.AssetId"
    Flag =2
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
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Freight"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Refrig"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Warranty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="Inspection"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="29"
    End
    Begin
        dbText "Name" ="Other"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="33"
    End
    Begin
        dbText "Name" ="PTO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="37"
    End
    Begin
        dbText "Name" ="CoNameDecal"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="40"
    End
    Begin
        dbText "Name" ="CoNameCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="41"
    End
    Begin
        dbText "Name" ="RegDecal"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="44"
    End
    Begin
        dbText "Name" ="RegDecalCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="45"
    End
    Begin
        dbText "Name" ="MiscDecal"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="48"
    End
    Begin
        dbText "Name" ="MiscCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="49"
    End
    Begin
        dbText "Name" ="TankBody"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="53"
    End
    Begin
        dbText "Name" ="TankOther"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="57"
    End
    Begin
        dbText "Name" ="OilRig"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="61"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Cost_FeesTax.TaxName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Cost_FeesTax.TaxInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Cost_FeesTax.TaxPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Cost_Freight.FreightName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Cost_Freight.FGTInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="Cost_Freight.FreightPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="Cost_Warranty.WarrantyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="Cost_Warranty.WarrantyInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="Cost_Warranty.WarrantyPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="Cost_Inspection.InspCo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="Cost_Inspection.InspInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="30"
    End
    Begin
        dbText "Name" ="Cost_Inspection.InspPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="31"
    End
    Begin
        dbText "Name" ="Cost_Other.OtherMake"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="Cost_Other.OtherInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="34"
    End
    Begin
        dbText "Name" ="Cost_Other.OtherPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="35"
    End
    Begin
        dbText "Name" ="Cost_PTO.PTOMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="36"
    End
    Begin
        dbText "Name" ="Cost_PTO.PTOInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="38"
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
        dbText "Name" ="Cost_TankOther.TankOthPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="59"
    End
    Begin
        dbText "Name" ="Cost_Oil.OilRigMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="60"
    End
    Begin
        dbText "Name" ="Cost_Oil.OilRigInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="62"
    End
    Begin
        dbText "Name" ="Cost_Oil.OilRigPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="63"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInstallPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="83"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Make"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="64"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Cost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="65"
    End
    Begin
        dbText "Name" ="TotalCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="66"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="67"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Adapt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="68"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCAdaptr"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="69"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCAdaptInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="70"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCAdaptPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="71"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Dome"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="72"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCDome"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="73"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCDomeInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="74"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCDomePaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="75"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Multi"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="76"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCMulti"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="77"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCMultiInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="78"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCMultiPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="79"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInstaller"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="80"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInstall"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="81"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInsInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="82"
    End
    Begin
        dbText "Name" ="RefrigMake"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigSerial"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.UnitCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="UnitOEC1.[Title Num]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="UnitOEC1.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="UnitOEC1.TitledState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.StreetAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.InService"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.AssetId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="UnitOEC1.Unitvin"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="UnitOEC1.Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="UnitOEC1.UnitInvoicePaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="UnitOEC1.Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.InterimBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.InterimDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.InterimAmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.unitno"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="UnitOEC1.GroupREF"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="UnitOEC1.DrawDownId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.VendorREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.LORent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC1.perdiemrent"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-378
    Top =25
    Right =1111
    Bottom =1031
    Left =-1
    Top =-1
    Right =1457
    Bottom =356
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =37
        Top =56
        Right =186
        Bottom =310
        Top =0
        Name ="UnitOEC1"
        Name =""
    End
    Begin
        Left =261
        Top =1
        Right =359
        Bottom =145
        Top =0
        Name ="Cost_FeesTax"
        Name =""
    End
    Begin
        Left =383
        Top =1
        Right =483
        Bottom =145
        Top =0
        Name ="Cost_Freight"
        Name =""
    End
    Begin
        Left =907
        Top =164
        Right =1033
        Bottom =308
        Top =0
        Name ="Cost_Refrig"
        Name =""
    End
    Begin
        Left =1061
        Top =163
        Right =1167
        Bottom =307
        Top =0
        Name ="Cost_Warranty"
        Name =""
    End
    Begin
        Left =511
        Top =4
        Right =655
        Bottom =148
        Top =0
        Name ="Cost_Other"
        Name =""
    End
    Begin
        Left =689
        Top =8
        Right =833
        Bottom =152
        Top =0
        Name ="Cost_PTO"
        Name =""
    End
    Begin
        Left =860
        Top =5
        Right =1004
        Bottom =149
        Top =0
        Name ="Cost_Inspection"
        Name =""
    End
    Begin
        Left =1029
        Top =8
        Right =1173
        Bottom =152
        Top =0
        Name ="Cost_DecalsCoName"
        Name =""
    End
    Begin
        Left =1382
        Top =12
        Right =1526
        Bottom =156
        Top =0
        Name ="Cost_DecalsRegNumber"
        Name =""
    End
    Begin
        Left =1184
        Top =12
        Right =1328
        Bottom =156
        Top =0
        Name ="Cost_TankBody"
        Name =""
    End
    Begin
        Left =1318
        Top =156
        Right =1462
        Bottom =300
        Top =0
        Name ="Cost_TankOther"
        Name =""
    End
    Begin
        Left =272
        Top =250
        Right =416
        Bottom =394
        Top =0
        Name ="Cost_Oil"
        Name =""
    End
    Begin
        Left =477
        Top =254
        Right =621
        Bottom =398
        Top =0
        Name ="Cost_DecalsMisc"
        Name =""
    End
    Begin
        Left =629
        Top =212
        Right =800
        Bottom =480
        Top =0
        Name ="Cost_OBC_All"
        Name =""
    End
End
