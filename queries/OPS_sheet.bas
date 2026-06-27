dbMemo "SQL" ="SELECT vw_OPS_Report_Base.Lessee, vw_OPS_Report_Base.[Lease No], vw_OPS_Report_B"
    "ase.SchNo, vw_OPS_Report_Base.[Group], vw_OPS_Report_Base.[Unit No], vw_OPS_Repo"
    "rt_Base.AssetId, vw_OPS_Report_Base.Status, vw_OPS_Report_Base.UnitVIN, vw_OPS_R"
    "eport_Base.[Desc], vw_OPS_Report_Base.Garage, vw_OPS_Report_Base.StreetAddress, "
    "vw_OPS_Report_Base.City, vw_OPS_Report_Base.State, vw_OPS_Report_Base.Zip, vw_OP"
    "S_Report_Base.County, vw_OPS_Report_Base.RegLoc, cRefrig.[Refrig Make], cRefrig."
    "[Refrig Class Type], cRefrig.[Refrig Serial], vw_OPS_Report_Base.PlateNum, vw_OP"
    "S_Report_Base.Delivered, vw_OPS_Report_Base.Accepted, vw_OPS_Report_Base.InServi"
    "ceInd, vw_OPS_Report_Base.InServiceDate, vw_OPS_Report_Base.MSORecv, vw_OPS_Repo"
    "rt_Base.MSOCopy, vw_OPS_Report_Base.TitleAppSent, vw_OPS_Report_Base.LastTitleLo"
    "g, vw_OPS_Report_Base.TitleAction, vw_OPS_Report_Base.UnitCost, vw_OPS_Report_Ba"
    "se.ChildPartCost, vw_OPS_Report_Base.UnitTotalCost, vw_OPS_Report_Base.VendorCos"
    "t, vw_OPS_Report_Base.Residual, vw_OPS_Report_Base.Rent, vw_OPS_Report_Base.Synd"
    "Rent, vw_OPS_Report_Base.PerDiemStart, vw_OPS_Report_Base.BLCD, vw_OPS_Report_Ba"
    "se.ProjEOL, vw_OPS_Report_Base.ODORead, vw_OPS_Report_Base.ODOSource, vw_OPS_Rep"
    "ort_Base.DeliveryLoc, vw_OPS_Report_Base.Title, vw_OPS_Report_Base.TitleState, v"
    "w_OPS_Report_Base.LienRelease, vw_OPS_Report_Base.CustodianID, vw_OPS_Report_Bas"
    "e.SchID\015\012FROM (vw_OPS_Report_Base LEFT JOIN Ops_ChildPart_Make ON Ops_Chil"
    "dPart_Make.AssetID = vw_OPS_Report_Base.AssetId) LEFT JOIN (SELECT\015\012      "
    "      UnitRef AS cAssetID,\015\012            ChildMake AS [Refrig Make],\015\012"
    "            ChildClass & ' ' & ChildType AS [Refrig Class Type],\015\012        "
    "    ChildVIN AS [Refrig Serial]         \015\012            FROM\015\012        "
    "    ChildParts\015\012        WHERE\015\012            ChildClass LIKE '*Temp*'\015"
    "\012            AND ChildClass LIKE '*Refrig*'\015\012    )  AS cRefrig ON cRefr"
    "ig.cAssetID = vw_OPS_Report_Base.AssetId;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbBoolean "NoFormat" ="0"
Begin
    Begin
        dbText "Name" ="InServiceInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Delivered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RefrigCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Accepted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.[Lease No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MSORecv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ops_ChildPart_Serial.[Refrig Unit_Serial]"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="titleAppSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleAction"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Location"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastTitleLog"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastMilesReportedAll.ODORead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MSOCopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Residual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ops_ChildPart_Make.[Refrigerated Truck Body_Make]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PerDiemStart"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastMilesReportedAll.LastMiles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cost_Refrig.RefrigSerialNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildPart_Cost_OEC_summed.ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ops_ChildPart_Make.[Refrig Body_Make]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.[Unit No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1016"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.StreetAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitTotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DeliveryLoc"
        dbInteger "ColumnWidth" ="3150"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ops_ChildPart_Make.[Refrig Unit_Make]"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ops_ChildPart_Make.[Refrigerated Body_Make]"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ops_ChildPart_Serial.[Refrigerated Body_Serial]"
        dbInteger "ColumnWidth" ="2625"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ops_ChildPart_Serial.[Refrigerated Truck Body_Serial]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_SixKeys.ProjEOL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LastMilesReportedAll.ODOSource"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitActualDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Garage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ops_ChildPart_Serial.[Refrig Body_Serial]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1017"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Title"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TitleState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTitleState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScheduleUnitDocs3.Vendorcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CustodianTitleTrustReLeasedUnits.nAssetID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustodianID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.CustodianID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.PlateNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.InServiceInd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.SchId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.lienrelease"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.[Lease No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.[Unit No]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ops_ChildPart_Make.Refrig_Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.Garage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ops_ChildPart_Serial.Refrig_Serial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.AssetId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.StreetAddress"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.County"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.ODOSource"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.RegLoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.Delivered"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.Accepted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.InServiceDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.MSORecv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.MSOCopy"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.TitleAppSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.LastTitleLog"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.TitleAction"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.UnitTotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.VendorCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.Residual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.SyndRent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.PerDiemStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.ProjEOL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.ODORead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.DeliveryLoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.Title"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.TitleState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.[Desc]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_OPS_Report_Base.[Group]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cRefrig.Refrig_Make"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cRefrig.Refrig_ClassType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cRefrig.Refrig_Serial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cRefrig.[Refrig Make]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cRefrig.[Refrig Class Type]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cRefrig.[Refrig Serial]"
        dbLong "AggregateType" ="-1"
    End
End
