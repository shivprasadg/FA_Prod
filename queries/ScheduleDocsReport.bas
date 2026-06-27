Operation =1
Option =0
Where ="(((ScheduleUnitDocs3.AssetId) Is Not Null))"
Begin InputTables
    Name ="ScheduleUnitDocs3"
    Name ="Cost_ChildParts_AllTypes"
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
    Alias ="Expr1"
    Expression ="Cost_ChildParts_AllTypes.[APU Make]"
    Alias ="Expr2"
    Expression ="Cost_ChildParts_AllTypes.[APU Cost]"
    Alias ="Expr3"
    Expression ="Cost_ChildParts_AllTypes.[APU Invoice]"
    Alias ="Expr4"
    Expression ="Cost_ChildParts_AllTypes.[APU Paid]"
    Alias ="Expr5"
    Expression ="Cost_ChildParts_AllTypes.[APU DrawNum]"
    Alias ="Expr6"
    Expression ="Cost_ChildParts_AllTypes.[Deer Guards Make]"
    Alias ="Expr7"
    Expression ="Cost_ChildParts_AllTypes.[Deer Guards Cost]"
    Alias ="Expr8"
    Expression ="Cost_ChildParts_AllTypes.[Deer Guards Invoice]"
    Alias ="Expr9"
    Expression ="Cost_ChildParts_AllTypes.[Deer Guards Paid]"
    Alias ="Expr10"
    Expression ="Cost_ChildParts_AllTypes.[Deer Guards DrawNum]"
    Alias ="Expr11"
    Expression ="Cost_ChildParts_AllTypes.[Headboards & Toolboxes Make]"
    Alias ="Expr12"
    Expression ="Cost_ChildParts_AllTypes.[Headboards & Toolboxes Cost]"
    Alias ="Expr13"
    Expression ="Cost_ChildParts_AllTypes.[Headboards & Toolboxes Invoice]"
    Alias ="Expr14"
    Expression ="Cost_ChildParts_AllTypes.[Headboards & Toolboxes Paid]"
    Alias ="Expr15"
    Expression ="Cost_ChildParts_AllTypes.[Headboards & Toolboxes DrawNum]"
    Alias ="Expr16"
    Expression ="Cost_ChildParts_AllTypes.[Installation Make]"
    Alias ="Expr17"
    Expression ="Cost_ChildParts_AllTypes.[Installation Cost]"
    Alias ="Expr18"
    Expression ="Cost_ChildParts_AllTypes.[Installation Invoice]"
    Alias ="Expr19"
    Expression ="Cost_ChildParts_AllTypes.[Installation Paid]"
    Alias ="Expr20"
    Expression ="Cost_ChildParts_AllTypes.[Installation DrawNum]"
    Alias ="Expr21"
    Expression ="Cost_ChildParts_AllTypes.[Inverters Make]"
    Alias ="Expr22"
    Expression ="Cost_ChildParts_AllTypes.[Inverters Cost]"
    Alias ="Expr23"
    Expression ="Cost_ChildParts_AllTypes.[Inverters Invoice]"
    Alias ="Expr24"
    Expression ="Cost_ChildParts_AllTypes.[Inverters Paid]"
    Alias ="Expr25"
    Expression ="Cost_ChildParts_AllTypes.[Inverters DrawNum]"
    Expression ="Cost_ChildParts_AllTypes.ChildPartTotal"
    Alias ="UnitTotalCost"
    Expression ="[UnitCost]+[ChildPartTotal]"
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
    RightTable ="Cost_ChildParts_AllTypes"
    Expression ="ScheduleUnitDocs3.AssetId = Cost_ChildParts_AllTypes.AssetId"
    Flag =1
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
        dbInteger "ColumnOrder" ="43"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.StreetAddress"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="44"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.City"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="45"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.State"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="46"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Zip"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="47"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Title Num"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="61"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="40"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.InService"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="41"
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
        dbInteger "ColumnOrder" ="42"
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
        dbInteger "ColumnOrder" ="48"
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
        dbInteger "ColumnOrder" ="39"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Inverters Invoice]"
        dbInteger "ColumnOrder" ="35"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Inverters Paid]"
        dbInteger "ColumnOrder" ="36"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Inverters DrawNum]"
        dbInteger "ColumnOrder" ="37"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Headboards & Toolboxes Invoice]"
        dbInteger "ColumnOrder" ="25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.ChildPartTotal"
        dbInteger "ColumnOrder" ="38"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Installation Make]"
        dbInteger "ColumnOrder" ="28"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Installation Cost]"
        dbInteger "ColumnOrder" ="29"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Deer Guards Paid]"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Inverters Make]"
        dbInteger "ColumnOrder" ="33"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Installation Invoice]"
        dbInteger "ColumnOrder" ="30"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[APU Cost]"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Deer Guards DrawNum]"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Inverters Cost]"
        dbInteger "ColumnOrder" ="34"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[APU Paid]"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Headboards & Toolboxes Paid]"
        dbInteger "ColumnOrder" ="26"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Deer Guards Make]"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Headboards & Toolboxes DrawNum]"
        dbInteger "ColumnOrder" ="27"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[APU Invoice]"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Deer Guards Cost]"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Installation Paid]"
        dbInteger "ColumnOrder" ="31"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[APU Make]"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Headboards & Toolboxes Cost]"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[APU DrawNum]"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Deer Guards Invoice]"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Installation DrawNum]"
        dbInteger "ColumnOrder" ="32"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_ChildParts_AllTypes.[Headboards & Toolboxes Make]"
        dbInteger "ColumnOrder" ="23"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1617
    Bottom =991
    Left =-1
    Top =-1
    Right =1585
    Bottom =218
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =28
        Top =8
        Right =194
        Bottom =352
        Top =0
        Name ="ScheduleUnitDocs3"
        Name =""
    End
    Begin
        Left =281
        Top =37
        Right =528
        Bottom =221
        Top =0
        Name ="Cost_ChildParts_AllTypes"
        Name =""
    End
End
