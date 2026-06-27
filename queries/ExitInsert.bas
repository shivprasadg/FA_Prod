dbMemo "SQL" ="INSERT INTO UnitExitOption ( UnitID, UnitGroupID, ProfileBLCD, ProfileTerm, Prof"
    "ileEOLDate, ProfileType, RentLO, RentPerdiem, RentSynd, RentActive, ResidualSynd"
    ", ResidualLO, AssigneeBankId, RentMTM, NoticeReqMonths, UnitOEC, RentExtended, E"
    "xchangeITAvailMths, RentHoldOver, RentIncrement, UnitLRF )\015\012SELECT Units.U"
    "nitID, Units.GroupREF AS UnitGroupID, SchGrp.BLCD AS OrigBLCD, SchGrp.LeaseTermP"
    "rimary AS Term, DateAdd(\"m\",[SchGrp].[LeaseTermPrimary],[SchGrp].[BLCD])-1 AS "
    "ProfileEOLDate, \"Primary Lease\" AS ProfileType, Units.UNITRENT AS RentLO, Unit"
    "s.UNITPDRENT AS RentPerdiem, Units.UNITSRENT AS RentSynd, Units.UNITSRENT AS Ren"
    "tActive, Units.UnitSynResid AS ResidualSynd, Units.UnitSalesResid, Banks.BankId,"
    " IIf(Nz([MTMRate],0)=0,[UnitSRent],CCur([Totalcost]*[MTMRate])) AS RentMTM, SchG"
    "rp.LeaseTermNotice AS NoticeReqMonths, UnitOEC2.TotalCost AS UnitOEC, CCur([Tota"
    "lCost]*[LeaseExtOption1LRF]) AS RentExtended, SchGrp.ExchangeITMonth AS Exchange"
    "ITAvailMths, Units.UNITSRENT AS RentHoldOver, 'Monthly' AS RentIncrement, IIf([U"
    "nitsRent]=0 Or [TotalCost]=0,0,Round([UnitsRent]/[TotalCost],6)) AS UnitLRF\015\012"
    "FROM ((((Units INNER JOIN SchGrp ON Units.GroupREF = SchGrp.SGrpID) INNER JOIN B"
    "anks ON SchGrp.Assignee_Bank = Banks.BankShortName) LEFT JOIN UnitExitOption ON "
    "Units.UnitID = UnitExitOption.UnitId) INNER JOIN vw_SixKeys ON Units.UnitID = vw"
    "_SixKeys.UnitID) INNER JOIN UnitOEC2 ON Units.UnitID = UnitOEC2.AssetId\015\012W"
    "HERE (((vw_SixKeys.MLOrig)=\"FA\"));\015\012"
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="BLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Termination"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Units.UnitStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Termanation"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitsrent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Units.UnitSynResid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Term"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="GroupOEC_LRF.GroupSyndLRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OptionStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Banks.BankId"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="UnitExitOption.UnitId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProfileEOLDate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="OrigBLCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="ProfileType"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="vw_Client_FAID_PO_Lease_Sch_Grp_Units.MLOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitpdrent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpO1LRF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.MTMRate"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpExtNot"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpExMo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Expr1007"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC2.TotalCost"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentExtended"
        dbInteger "ColumnWidth" ="1740"
        dbInteger "ColumnOrder" ="13"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentMTM"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Units.unitrent"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1008"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentSynd"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentActive"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitGroupID"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentLO"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentPerdiem"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProfileResidual"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NoticeReqMonths"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RentHoldOver"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ResidualSynd"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchangeITAvailMths"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitOEC"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UnitSalesResid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="20"
    End
    Begin
        dbText "Name" ="RentIncrement"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="UnitLRF"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
