dbMemo "SQL" ="SELECT Switch([SGRPPLB]=False,Null,Nz([PLBMileage],0)>0,Format([PLBMileage],\"#,"
    "##0\"),Nz([PLBHours],0),Format([PLBHours],\"#,##0\"),Nz([PLBMileage],0)>0 And Nz"
    "([PLBHours],0)>0,Format([PLBMileage],\"#,##0\") & \" / \" & Format([PLBHours],\""
    "#,##0\")) AS PLBMilesHours, Switch([SGRPPLB]=False,Null,Nz([PLBMileage],0)>0,\"M"
    "iles as of \" & Format([Forms]![BulkChange_Form]![PLBAsOfDate],\"mm/dd/yy\"),Nz("
    "[PLBHours],0)>0,\"Hours\",Nz([PLBMileage],0)=0 And Nz([PLBHours],0)=0,\"Miles / "
    "Hours\") AS PLBLabel, vw_CofABase.*\015\012FROM vw_CofABase INNER JOIN Bulk_Upda"
    "te_WorkTable ON vw_CofABase.UnitID = Bulk_Update_WorkTable.UnitID\015\012WHERE ("
    "((vw_CofABase.SGrpID)=[Forms]![BulkChange_Form]![GrpID]));\015\012"
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
Begin
    Begin
        dbText "Name" ="PLBLabel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PLBMilesHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitTitleLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.RegState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ChildVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SubLeasee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.DescriptionLong"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.LCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UNITRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ModelYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.FullNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.RegLoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Rentx"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.GrpCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Ctype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ChildPartID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Mileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Hours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.AmendmentNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitDeliveryLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ClientSignFullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.AmendmentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitTransDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitUnitNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.GarSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.LeaseTermPrimary"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.PLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ReadDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SGRPPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.CofASigner"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SchedDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientCoNm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.clientDBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Address1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Address2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Zip/Postal Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLMo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MLDay"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.CorpState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.GarageLoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitMake"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.MakeModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitSubType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.contactTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.BLCD_Projected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.UnitTax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.PLBMileage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.PLBHours"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.DeliveryLoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.TitleLoc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.TitledState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.DSub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.TaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ClientSignTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.RegSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.DelSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.TotalCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ChildYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ChildModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.ChildMakeModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.DCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Dst1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.Lessee"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vw_CofABase.DCNTY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bulk_Update_WorkTable.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bulk_Update_WorkTable.UnitDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bulk_Update_WorkTable.ChildPartCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bulk_Update_WorkTable.DrawDownID"
        dbLong "AggregateType" ="-1"
    End
End
