dbMemo "SQL" ="SELECT Schedule.SchID, Schedule.SchNo, SchGrp.SGrpNO AS GrpNo, ProjBLCDq.ProjBLC"
    "D, [schedule].[schno] & \"_\" & [Sgrpno] AS Sch, Schedule.SchShortDesc AS ShortD"
    "esc, SchGrp.LeaseTermPrimary AS PTErm, SchGrp.UnitsInGroup AS NoofUnits, SchGrp."
    "sgrpfslpmt AS FSLPmt, FSL.ID AS FSLSch, FSL.FSLYR1, FSL.FSLYR2, FSL.FSLYR3, FSL."
    "FSLYR4, FSL.FSLYR5, FSL.FSLYR6, FSL.FSLYR7, Clients.ClientShNm\015\012FROM ((((("
    "(Clients INNER JOIN MstrLease ON Clients.clientID=MstrLease.Client) INNER JOIN S"
    "chedule ON MstrLease.MLID=Schedule.SchMLRef) INNER JOIN SchGrp ON Schedule.SchID"
    "=SchGrp.SGrpSchID) INNER JOIN FSL ON Schedule.FSLSCH=FSL.ID) INNER JOIN Units ON"
    " SchGrp.SGrpID=Units.GroupREF) INNER JOIN Locations ON Units.UnitGarageLocationR"
    "ef=Locations.LocationID) LEFT JOIN ProjBLCDq ON Schedule.SchID=ProjBLCDq.SchID\015"
    "\012GROUP BY Schedule.SchID, Schedule.SchNo, SchGrp.SGrpNO, ProjBLCDq.ProjBLCD, "
    "[schedule].[schno] & \"_\" & [Sgrpno], Schedule.SchShortDesc, SchGrp.LeaseTermPr"
    "imary, SchGrp.UnitsInGroup, SchGrp.sgrpfslpmt, FSL.ID, FSL.FSLYR1, FSL.FSLYR2, F"
    "SL.FSLYR3, FSL.FSLYR4, FSL.FSLYR5, FSL.FSLYR6, FSL.FSLYR7, Clients.ClientShNm, S"
    "chGrp.FundDate\015\012HAVING (((Schedule.SchID)>24) AND ((Clients.ClientShNm)=\""
    "costco\"))\015\012ORDER BY Schedule.SchID, Schedule.SchNo, SchGrp.SGrpNO;\015\012"
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
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Schedule.SchNo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="FSL.FSLYR1"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="FSL.FSLYR2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="FSL.FSLYR3"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1050"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="FSL.FSLYR4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="FSL.FSLYR5"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="FSL.FSLYR6"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1050"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="FSL.FSLYR7"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1050"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="GrpNo"
        dbInteger "ColumnWidth" ="1020"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="ShortDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4290"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="PTErm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1005"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="NoofUnits"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="FSLPmt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="FSLSch"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1035"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="930"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Sch"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="ProjBLCDq.ProjBLCD"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
End
