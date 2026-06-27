dbMemo "SQL" ="SELECT MstrLease.MLNo, Schedule.SchID, Schedule.SchNo, [forms]![BillingSubForm]!"
    "[billfrom] AS STDt, [forms]![BillingSubForm]![billto] AS enddt, Units.unitunitnu"
    "m, Units.unitvin, [unityr] & \" \" & [unitmake] & \" \" & [unitmodel] & \" \" & "
    "[unittype] AS unitdesc, Units.unitvendorpytdate, SchGrp.UnitsInGroup, IIf([unitv"
    "endorpytdate]<[forms]![BillingSubForm]![billfrom],[forms]![BillingSubForm]![bill"
    "from],[unitvendorpytdate]) AS INtStDt, IIf([unitinservind]=\"a\" And [unitinserv"
    "]<[forms]![BillingSubForm]![billto],[unitinserv]-1,[forms]![BillingSubForm]![bil"
    "lto]) AS IntEndDt, SchGrp.SGRPINTRT, Round(DateDiff(\"d\",[unitvendorpytdate],II"
    "f([unitinservind]=\"a\" And [unitinserv]<[forms]![BillingSubForm]![billto],[unit"
    "inserv],[forms]![BillingSubForm]![billto]))*[unitfinalcost]*[SGRPINTRT]/360,2) A"
    "S IntAmt, Units.UnitInServ, Units.UnitInServInd, Units.UNITPDRENT, IIf(IsNull([u"
    "nitinservind]),Null,IIf([unitinserv]<[forms]![BillingSubForm]![billfrom],[forms]"
    "![BillingSubForm]![billfrom],[unitinserv])) AS PDStart, IIf(IsNull([unitinservin"
    "d]) Or [unitinservind]<>\"a\",Null,[forms]![BillingSubForm]![billto]) AS PDend, "
    "IIf([unitinservind]<>\"a\" Or IsNull([unitinservind]),-1,DateDiff(\"d\",IIf(IsNu"
    "ll([unitinservind]),Null,IIf([unitinserv]<[forms]![BillingSubForm]![billfrom],[f"
    "orms]![BillingSubForm]![billfrom],[unitinserv])),IIf(IsNull([unitinservind]) Or "
    "[unitinservind]<>\"a\",Null,[forms]![BillingSubForm]![billto])))+1 AS PDDAys, (I"
    "If([unitinservind]<>\"a\",0,DateDiff(\"d\",IIf(IsNull([unitinservind]),Null,IIf("
    "[unitinserv]<[forms]![BillingSubForm]![billfrom],[forms]![BillingSubForm]![billf"
    "rom],[unitinserv])),IIf(IsNull([unitinservind]) Or [unitinservind]<>\"a\",Null,["
    "forms]![BillingSubForm]![billto]))))*[unitpdrent] AS PDRent, Units.UnitFinalCost"
    ", [SGRPINTRT]/360 AS DailyIntRt, Round(DateDiff(\"d\",[unitvendorpytdate],IIf([u"
    "nitinservind]=\"a\" And [unitinserv]<[forms]![BillingSubForm]![billto],[unitinse"
    "rv],[forms]![BillingSubForm]![billto]))) AS IntDays, Locations.LocationStreet AS"
    " bloc, Locations.LAddrLine1 AS bloc2, Locations.LAddrLine2 AS BiLo3, Locations.L"
    "City AS BCity, Locations.LState AS Bst, Locations.LZip AS BZip, SchGrp.SGrpNO, L"
    "ocations_1.LocationStreet, Locations_1.LAddrLine2, Locations_1.LAddrLine3, Locat"
    "ions_1.LCity, Locations_1.LCounty, Locations_1.LState, Locations_1.LZip, Locatio"
    "ns_1.LTaxRate, Units.UnitTaxExempt, Locations_1.LocationName, SchGrp.SGRPINTRT, "
    "SchGrp.SGRPINTPD, SchGrp.SGRPNoPD, Clients.FAVenorNo, Right([unitvin],6) AS vin4"
    ", Schedule.SchFSL, lookup_States.StTaxType, UCase([invfld1desc] & \" \" & [invfl"
    "d1]) AS Bill1, UCase([invfld2desc] & \" \" & [invfld2]) AS Bill2, Bills.DueonRec"
    ", IIf([SGrpIntPD]=0,Null,[intdays]*[DailyIntRt]*[unitfinalcost])+(IIf([sgrpnopd]"
    "=Yes Or [pddays]=0,0,[pddays]*[unitpdrent]))+(IIf([SchFSL]=-1,Null,(IIf([unittax"
    "exempt]=Yes Or [sttaxtype]<>\"rr\",0,((IIf([SGrpIntPD]=0,Null,0))+(IIf([sgrpnopd"
    "]=Yes Or [pddays]=0,0,[pddays]*[unitpdrent])))*[locations_1].[LTaxRate])))) AS L"
    "ineTot\015\012FROM ((((Clients INNER JOIN Locations ON Clients.clientID=Location"
    "s.ClientREF) INNER JOIN MstrLease ON Clients.clientID=MstrLease.Client) INNER JO"
    "IN (Schedule LEFT JOIN Bills ON Schedule.SchMLRef=Bills.BillSchRef) ON MstrLease"
    ".MLID=Schedule.SchMLRef) INNER JOIN SchGrp ON Schedule.SchID=SchGrp.SGrpSchID) I"
    "NNER JOIN ((Units INNER JOIN Locations AS Locations_1 ON Units.UnitGarageLocatio"
    "nRef=Locations_1.LocationID) INNER JOIN lookup_States ON Locations_1.LState=look"
    "up_States.StateAbbrev) ON SchGrp.SGrpID=Units.GroupREF\015\012WHERE (((Schedule."
    "SchID)=forms!Billing!selsch) And ((Units.unitvendorpytdate) Is Not Null And (Uni"
    "ts.unitvendorpytdate)<=Forms!BillingSubForm!billto) And ((MstrLease.MLOrig)=\"fa"
    "\") And ((Units.unitstatus)=\"a\") And ((Locations.LocationType)=\"billing\"))\015"
    "\012ORDER BY Schedule.SchID, Schedule.SchNo, Units.unitunitnum;\015\012"
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
        dbText "Name" ="MstrLease.MLNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="STDt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="enddt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="Units.unitunitnum"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Units.unitvin"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="unitdesc"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Units.UnitVendorPytDate"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="PDDAys"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServInd"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PDRent"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IntAmt"
        dbInteger "ColumnWidth" ="1185"
        dbInteger "ColumnOrder" ="22"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitFinalCost"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNoUnits"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="23"
    End
    Begin
        dbText "Name" ="INtStDt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="IntEndDt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="PDStart"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="PDend"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="DailyIntRt"
        dbText "Format" ="Percent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
        dbByte "DecimalPlaces" ="6"
    End
    Begin
        dbText "Name" ="IntDays"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="bloc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="bloc2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BiLo3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bst"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocationStreet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocSt3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocSt4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocationCounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocationZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations_1.LocTaxRt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitTaxExempt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITPDRENT"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1605"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Locations_1.LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPINTRT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPINTPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPNoPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.FAVenorNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vin4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_States.StTaxType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bill1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bill2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Bills.DueonRec"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LineTot"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitInServ"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="24"
    End
    Begin
        dbText "Name" ="Expr1012"
        dbLong "AggregateType" ="-1"
    End
End
