Operation =1
Option =0
Begin InputTables
    Name ="CofA_Report_Mass_Pre"
    Name ="CofA_ChildParts"
End
Begin OutputColumns
    Expression ="CofA_Report_Mass_Pre.*"
    Expression ="CofA_ChildParts.ChildYear"
    Alias ="ChildMakeModel"
    Expression ="[ChildMake] & \" \" & [ChildClass]"
    Expression ="CofA_ChildParts.ChildVIN"
    Expression ="CofA_ChildParts.childtype"
    Alias ="Ctype"
    Expression ="IIf([childtype]=\"Refrig Unit\",\"Refrigeration Unit\",[ChildType])"
    Alias ="Lessee"
    Expression ="[clientCompanyName] & Chr(13) & Chr(10) & IIf(IsNull([clientCoNm2]),\"\",[client"
        "CoNm2] & Chr(13) & Chr(10)) & [clientAddress1] & \" \" & [clientAddress2] & Chr("
        "13) & Chr(10) & [clientcity] & \", \" & [clientstate] & \" \" & [clientzip]"
    Alias ="CofA_Date"
    Expression ="[WhichDate]"
End
Begin Joins
    LeftTable ="CofA_Report_Mass_Pre"
    RightTable ="CofA_ChildParts"
    Expression ="CofA_Report_Mass_Pre.UnitID = CofA_ChildParts.AssetId"
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
        dbText "Name" ="CofA_ChildParts.ChildYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ChildMakeModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_ChildParts.ChildVIN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UnitTitleLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UnitGarageLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UnitRegLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UnitDeliveryLocationRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.MstrLease.MLID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Schedule.SchedDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UnitID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.unitvin"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.GrpCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Clients.clientCoNm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Clients.clientDBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Clients.clientAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Clients.clientAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Clients.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Clients.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Clients.clientZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.MstrLease.MLDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.MLYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.MLMo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.mlday"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Schedule.SchNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.CorpState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.PRINTCOFA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.ModelYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UNITMAKE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.MakeModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UnitType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.Axle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Desc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.unitunitnum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UnitAcceptDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.Unitstatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.SchGrp.SGrpSchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.SchGrp.SGRPPLB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UnitTransDays"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.unitrent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.SchGrp.SGrpFSLPmt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UnitModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UNITTAX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Unitcost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CofA_Report_X.Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.TSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DSub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.Dst1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DST2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DCITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DCNTY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DzIP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.TaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.RSub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.Rst1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.Rst2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.RCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.RST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.RZIP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DelSub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DelSt1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DelSt2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DelCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DelSt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.DelZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.CLASS.TITLED"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.MLAmends.MLAmtNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.MLAmends.MLAmdDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.Contacts.contactTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.FullNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.MLRMo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.MLRday"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.MLRYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.Shmo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.Schday"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.SchYr"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.GarLOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.Locations_1.LocationCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.Locations_1.LocTaxRt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.ContactsIncumbent.ClientSignTitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.ContactsIncumbent.ClientSignFullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Report_Mass_Pre.UnitDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_ChildParts.ChildType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ctype"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lessee"
        dbInteger "ColumnWidth" ="6780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CofA_Date"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =95
    Top =93
    Right =2525
    Bottom =925
    Left =-1
    Top =-1
    Right =2398
    Bottom =413
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =52
        Top =25
        Right =283
        Bottom =550
        Top =0
        Name ="CofA_Report_Mass_Pre"
        Name =""
    End
    Begin
        Left =367
        Top =75
        Right =605
        Bottom =390
        Top =0
        Name ="CofA_ChildParts"
        Name =""
    End
End
