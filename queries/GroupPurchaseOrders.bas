dbMemo "SQL" ="SELECT SchGrp.POVendor, SchGrp.POSentDt As POSentDate, SchGrp.PONum1 As PONumber"
    ", 'UnitOEMPO' As POType, sGrpId As GroupId\015\012FROM SchGrp\015\012Union ALL\015"
    "\012SELECT PO2Vendor, PO2SentDt, PONum2, 'UpFitPO1' As POType, sGrpId As GroupId"
    "\015\012FROM SchGrp\015\012Union ALL\015\012SELECT PO3Vendor, PO3SentDt, PONum3,"
    " 'UpFitPO2' As POType, sGrpId As GroupId\015\012FROM SchGrp\015\012UNION ALL SEL"
    "ECT PO4Vendor, PO4SentDt, PONum4, 'UpFitPO3' As POType, sGrpId As GroupId\015\012"
    "FROM SchGrp;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "Filter" ="([Query1].[GroupId]=70)"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="SchGrp.POVendor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="POType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PONumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="POSentDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupId"
        dbLong "AggregateType" ="-1"
    End
End
