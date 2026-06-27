Operation =1
Option =0
Where ="(((LeaseReport1.UnitType) Not Like \"*Warranty*\"))"
Begin InputTables
    Name ="LeaseReport1"
    Name ="Cost_FeesTax"
    Name ="Cost_Freight"
    Name ="Cost_Refrig"
    Name ="Cost_Warranty"
    Name ="Cost_Other"
    Name ="Cost_PTO"
    Name ="Cost_Inspection"
    Name ="Cost_DecalsCoName"
    Name ="Cost_DecalsRegNumber"
    Name ="Cost_DecalsMisc"
    Name ="Cost_TankBody"
    Name ="Cost_TankOther"
    Name ="Cost_Oil"
    Name ="Cost_OBC_All"
End
Begin OutputColumns
    Expression ="LeaseReport1.[Lease No]"
    Expression ="LeaseReport1.Lessee"
    Expression ="LeaseReport1.Co_Lessee"
    Alias ="SYDYear"
    Expression ="IIf(IsDate([SynDate]),DatePart(\"yyyy\",[SynDate]),0)"
    Alias ="OrderedYear"
    Expression ="IIf([Status]=\"B\" And Not IsDate([UnitInvoicePaid]),DatePart(\"yyyy\",Nz([UnitE"
        "stDelDate],Nz([PODate],[UnitCOABack]))),0)"
    Alias ="OwnedYear"
    Expression ="IIf([Status]=\"B\" And IsDate([UnitInvoicePaid]),DatePart(\"yyyy\",[UnitInvoiceP"
        "aid]),0)"
    Alias ="LeaseYear"
    Expression ="IIf(IsDate([SynDate]),DatePart(\"yyyy\",[SynDate]),0)"
    Alias ="InServYear"
    Expression ="IIf(IsNull([unitcoaback]),[OrderedYear],CDbl(DatePart(\"yyyy\",[unitcoaback])))"
    Expression ="LeaseReport1.FAID"
    Alias ="LeaseMonth"
    Expression ="DatePart(\"m\",[MstrLeaseDate])"
    Expression ="LeaseReport1.SchNo"
    Expression ="LeaseReport1.ModelYear"
    Expression ="LeaseReport1.Make"
    Expression ="LeaseReport1.UnitType"
    Expression ="LeaseReport1.Group"
    Expression ="LeaseReport1.[Unit No]"
    Expression ="LeaseReport1.Unitvin"
    Expression ="LeaseReport1.Desc"
    Expression ="LeaseReport1.UnitCost"
    Expression ="LeaseReport1.UnitInvoiceNum"
    Expression ="LeaseReport1.UnitInvoicePaid"
    Alias ="RefrigMake_Serial"
    Expression ="IIf(IsNull([RefrigMake]),Null,[RefrigMake] & \" - \" & [RefrigSerialNum])"
    Alias ="Refrig"
    Expression ="CCur(Nz([Cost_Refrig].[Refrig],0))"
    Expression ="Cost_Refrig.RefrigInvoice"
    Expression ="Cost_Refrig.RefrigPaid"
    Alias ="OBCMake"
    Expression ="Cost_OBC_All.OBC_Make"
    Alias ="OBC"
    Expression ="CCur(Nz([Cost_OBC_All].[OBC_Cost],0))"
    Expression ="Cost_OBC_All.OBCInvoice"
    Expression ="Cost_OBC_All.OBCPaid"
    Expression ="Cost_OBC_All.OBC_Adapt"
    Alias ="OBCAdapt"
    Expression ="CCur(Nz([Cost_OBC_All].[OBCAdaptr],0))"
    Expression ="Cost_OBC_All.OBCAdaptInv"
    Expression ="Cost_OBC_All.OBCAdaptPaid"
    Expression ="Cost_OBC_All.OBC_Dome"
    Alias ="OBCDome"
    Expression ="CCur(Nz([Cost_OBC_All].[OBCDome],0))"
    Expression ="Cost_OBC_All.OBCDomeInv"
    Expression ="Cost_OBC_All.OBCDomePaid"
    Expression ="Cost_OBC_All.OBC_Multi"
    Alias ="OBCMulti"
    Expression ="CCur(Nz([Cost_OBC_All].[OBCMulti],0))"
    Expression ="Cost_OBC_All.OBCMultiInv"
    Expression ="Cost_OBC_All.OBCMultiPaid"
    Expression ="Cost_OBC_All.OBCInstaller"
    Alias ="OBCInstall"
    Expression ="CCur(Nz([Cost_OBC_All].[OBCInstall],0))"
    Expression ="Cost_OBC_All.OBCInsInvoice"
    Expression ="Cost_OBC_All.OBCInstallPaid"
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
    Alias ="TankBody"
    Expression ="CCur(Nz([Cost_TankBody].[tankbody],0))"
    Expression ="Cost_TankBody.TankMake"
    Expression ="Cost_TankBody.TankInvoice"
    Expression ="Cost_TankBody.TankPaid"
    Alias ="TankOther"
    Expression ="CCur(Nz([Cost_TankOther].[tankOther],0))"
    Expression ="Cost_TankOther.TankOthMake"
    Expression ="Cost_TankOther.TankOthInvoice"
    Expression ="Cost_TankOther.TankOthPaid"
    Alias ="OilRig"
    Expression ="CCur(Nz([Cost_Oil].[OilRig],0))"
    Expression ="Cost_Oil.OilRigMake"
    Expression ="Cost_Oil.OilRigInvoice"
    Expression ="Cost_Oil.OilRigPaid"
    Alias ="TotalChildCost"
    Expression ="CCur([Refrig]+[OBCTotal]+[FeesTax]+[Freight]+[Warranty]+[Other]+[Inspection]+[PT"
        "O]+[CoNameCost]+[RegDecalCost]+[MiscCost]+[TankBody]+[TankOther]+[OilRig])"
    Alias ="TotalCost"
    Expression ="CCur([UnitCost]+[Refrig]+[OBCTotal]+[FeesTax]+[Freight]+[Warranty]+[Other]+[Insp"
        "ection]+[PTO]+[CoNameCost]+[RegDecalCost]+[MiscCost]+[TankBody]+[TankOther]+[Oil"
        "Rig])"
    Expression ="LeaseReport1.AssetId"
    Expression ="LeaseReport1.Status"
    Expression ="LeaseReport1.InService"
    Expression ="LeaseReport1.[Title Num]"
    Alias ="Ordered"
    Expression ="IIf(Not IsDate([UnitInvoicePaid]) And Not IsDate([SynDate]),\"Ordered\",Null)"
    Alias ="Owned"
    Expression ="IIf(IsNull([SynDate]) And IsDate([UnitInvoicePaid]),\"Owned\",Null)"
    Alias ="Syndicated"
    Expression ="IIf(IsDate([SynDate]),\"Syndicated\",Null)"
End
Begin Joins
    LeftTable ="LeaseReport1"
    RightTable ="Cost_FeesTax"
    Expression ="LeaseReport1.AssetId = Cost_FeesTax.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_Freight"
    Expression ="LeaseReport1.AssetId = Cost_Freight.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_Refrig"
    Expression ="LeaseReport1.AssetId = Cost_Refrig.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_Warranty"
    Expression ="LeaseReport1.AssetId = Cost_Warranty.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_Other"
    Expression ="LeaseReport1.AssetId = Cost_Other.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_PTO"
    Expression ="LeaseReport1.AssetId = Cost_PTO.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_Inspection"
    Expression ="LeaseReport1.AssetId = Cost_Inspection.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_DecalsCoName"
    Expression ="LeaseReport1.AssetId = Cost_DecalsCoName.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_DecalsRegNumber"
    Expression ="LeaseReport1.AssetId = Cost_DecalsRegNumber.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_DecalsMisc"
    Expression ="LeaseReport1.AssetId = Cost_DecalsMisc.DecalVendor"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_TankBody"
    Expression ="LeaseReport1.AssetId = Cost_TankBody.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_TankOther"
    Expression ="LeaseReport1.AssetId = Cost_TankOther.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_Oil"
    Expression ="LeaseReport1.AssetId = Cost_Oil.AssetId"
    Flag =2
    LeftTable ="LeaseReport1"
    RightTable ="Cost_OBC_All"
    Expression ="LeaseReport1.AssetId = Cost_OBC_All.AssetId"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[LeaseReport2].[InServYear]"
Begin
    Begin
        dbText "Name" ="RefrigMake_Serial"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="Refrig"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="OBCMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="27"
    End
    Begin
        dbText "Name" ="OBC"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="28"
    End
    Begin
        dbText "Name" ="OBCInstall"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="32"
    End
    Begin
        dbText "Name" ="FeesTax"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="36"
    End
    Begin
        dbText "Name" ="Freight"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="40"
    End
    Begin
        dbText "Name" ="Warranty"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="44"
    End
    Begin
        dbText "Name" ="Inspection"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="48"
    End
    Begin
        dbText "Name" ="Other"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="52"
    End
    Begin
        dbText "Name" ="PTO"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="56"
    End
    Begin
        dbText "Name" ="CoNameDecal"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="59"
    End
    Begin
        dbText "Name" ="CoNameCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="60"
    End
    Begin
        dbText "Name" ="RegDecal"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="63"
    End
    Begin
        dbText "Name" ="RegDecalCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="64"
    End
    Begin
        dbText "Name" ="MiscDecal"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="67"
    End
    Begin
        dbText "Name" ="MiscCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="68"
    End
    Begin
        dbText "Name" ="TankBody"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="71"
    End
    Begin
        dbText "Name" ="TankOther"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="75"
    End
    Begin
        dbText "Name" ="OilRig"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="79"
    End
    Begin
        dbText "Name" ="TotalCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="84"
    End
    Begin
        dbText "Name" ="LeaseReport1.[Lease No]"
        dbInteger "ColumnWidth" ="4110"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="LeaseReport1.Lessee"
        dbInteger "ColumnWidth" ="4110"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="LeaseReport1.Co_Lessee"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="LeaseReport1.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
        dbInteger "ColumnWidth" ="1530"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="LeaseReport1.Group"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="LeaseReport1.[Unit No]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="LeaseReport1.Unitvin"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="19"
    End
    Begin
        dbText "Name" ="LeaseReport1.Desc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="LeaseReport1.UnitCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="21"
    End
    Begin
        dbText "Name" ="LeaseReport1.UnitInvoiceNum"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="22"
    End
    Begin
        dbText "Name" ="LeaseReport1.UnitInvoicePaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="25"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="26"
    End
    Begin
        dbText "Name" ="Cost_FeesTax.TaxName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="35"
    End
    Begin
        dbText "Name" ="Cost_FeesTax.TaxInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="37"
    End
    Begin
        dbText "Name" ="Cost_FeesTax.TaxPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="38"
    End
    Begin
        dbText "Name" ="Cost_Freight.FreightName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="39"
    End
    Begin
        dbText "Name" ="Cost_Freight.FGTInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="41"
    End
    Begin
        dbText "Name" ="Cost_Freight.FreightPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="42"
    End
    Begin
        dbText "Name" ="Cost_Warranty.WarrantyName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="43"
    End
    Begin
        dbText "Name" ="Cost_Warranty.WarrantyInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="45"
    End
    Begin
        dbText "Name" ="Cost_Warranty.WarrantyPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="46"
    End
    Begin
        dbText "Name" ="Cost_Inspection.InspCo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="47"
    End
    Begin
        dbText "Name" ="Cost_Inspection.InspInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="49"
    End
    Begin
        dbText "Name" ="Cost_Inspection.InspPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="50"
    End
    Begin
        dbText "Name" ="Cost_Other.OtherMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="51"
    End
    Begin
        dbText "Name" ="Cost_Other.OtherInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="53"
    End
    Begin
        dbText "Name" ="Cost_Other.OtherPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="54"
    End
    Begin
        dbText "Name" ="Cost_PTO.PTOMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="55"
    End
    Begin
        dbText "Name" ="Cost_PTO.PTOInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="57"
    End
    Begin
        dbText "Name" ="Cost_PTO.PTOPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="58"
    End
    Begin
        dbText "Name" ="Cost_DecalsCoName.CoNameInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="61"
    End
    Begin
        dbText "Name" ="Cost_DecalsCoName.CoNamePaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="62"
    End
    Begin
        dbText "Name" ="Cost_DecalsRegNumber.RegDecalInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="65"
    End
    Begin
        dbText "Name" ="Cost_DecalsRegNumber.RegDecalPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="66"
    End
    Begin
        dbText "Name" ="Cost_DecalsMisc.MiscInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="69"
    End
    Begin
        dbText "Name" ="Cost_DecalsMisc.MiscPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="70"
    End
    Begin
        dbText "Name" ="Cost_TankBody.TankMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="72"
    End
    Begin
        dbText "Name" ="Cost_TankBody.TankInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="73"
    End
    Begin
        dbText "Name" ="Cost_TankBody.TankPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="74"
    End
    Begin
        dbText "Name" ="Cost_TankOther.TankOthMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="76"
    End
    Begin
        dbText "Name" ="Cost_TankOther.TankOthInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="77"
    End
    Begin
        dbText "Name" ="Cost_TankOther.TankOthPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="78"
    End
    Begin
        dbText "Name" ="Cost_Oil.OilRigMake"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="80"
    End
    Begin
        dbText "Name" ="Cost_Oil.OilRigInvoice"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="81"
    End
    Begin
        dbText "Name" ="Cost_Oil.OilRigPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="82"
    End
    Begin
        dbText "Name" ="LeaseReport1.AssetId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="97"
    End
    Begin
        dbText "Name" ="LeaseReport1.Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="98"
    End
    Begin
        dbText "Name" ="LeaseReport1.InService"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="99"
    End
    Begin
        dbText "Name" ="LeaseReport1.[Title Num]"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="100"
    End
    Begin
        dbText "Name" ="LeaseYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="LeaseReport1.FAID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="LeaseMonth"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="LeaseReport1.Make"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="LeaseReport1.UnitType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="TotalChildCost"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="83"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInsInvoice"
        dbInteger "ColumnOrder" ="33"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInstallPaid"
        dbInteger "ColumnOrder" ="34"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInstaller"
        dbInteger "ColumnOrder" ="31"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCPaid"
        dbInteger "ColumnOrder" ="30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCInvoice"
        dbInteger "ColumnOrder" ="29"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCMultiPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="96"
    End
    Begin
        dbText "Name" ="OBCAdapt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="86"
    End
    Begin
        dbText "Name" ="OBCDome"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="90"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Dome"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="89"
    End
    Begin
        dbText "Name" ="OBCMulti"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="94"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Adapt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="85"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCAdaptInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="87"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCAdaptPaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="88"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCDomeInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="91"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCDomePaid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="92"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBC_Multi"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="93"
    End
    Begin
        dbText "Name" ="Cost_OBC_All.OBCMultiInv"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="95"
    End
    Begin
        dbText "Name" ="OrderedYear"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SYDYear"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="Owned"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="Ordered"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Syndicated"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="InServYear"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OwnedYear"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseReport1.ModelYear"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =174
    Top =6
    Right =1832
    Bottom =974
    Left =-1
    Top =-1
    Right =1626
    Bottom =317
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =61
        Top =14
        Right =205
        Bottom =389
        Top =0
        Name ="LeaseReport1"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Cost_FeesTax"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="Cost_Freight"
        Name =""
    End
    Begin
        Left =818
        Top =10
        Right =962
        Bottom =154
        Top =0
        Name ="Cost_Refrig"
        Name =""
    End
    Begin
        Left =240
        Top =309
        Right =384
        Bottom =453
        Top =0
        Name ="Cost_Warranty"
        Name =""
    End
    Begin
        Left =240
        Top =156
        Right =384
        Bottom =300
        Top =0
        Name ="Cost_Other"
        Name =""
    End
    Begin
        Left =432
        Top =156
        Right =576
        Bottom =300
        Top =0
        Name ="Cost_PTO"
        Name =""
    End
    Begin
        Left =624
        Top =156
        Right =768
        Bottom =300
        Top =0
        Name ="Cost_Inspection"
        Name =""
    End
    Begin
        Left =816
        Top =156
        Right =960
        Bottom =300
        Top =0
        Name ="Cost_DecalsCoName"
        Name =""
    End
    Begin
        Left =1008
        Top =156
        Right =1152
        Bottom =300
        Top =0
        Name ="Cost_DecalsRegNumber"
        Name =""
    End
    Begin
        Left =436
        Top =306
        Right =580
        Bottom =450
        Top =0
        Name ="Cost_DecalsMisc"
        Name =""
    End
    Begin
        Left =622
        Top =314
        Right =766
        Bottom =458
        Top =0
        Name ="Cost_TankBody"
        Name =""
    End
    Begin
        Left =816
        Top =310
        Right =960
        Bottom =454
        Top =0
        Name ="Cost_TankOther"
        Name =""
    End
    Begin
        Left =1010
        Top =308
        Right =1154
        Bottom =452
        Top =0
        Name ="Cost_Oil"
        Name =""
    End
    Begin
        Left =627
        Top =15
        Right =770
        Bottom =152
        Top =0
        Name ="Cost_OBC_All"
        Name =""
    End
End
