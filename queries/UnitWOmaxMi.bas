dbMemo "SQL" ="SELECT [costco repairsNew].Unitref, Max([costco repairsNew].HubMI) AS MaxOfHubMI"
    ", Max([costco repairsNew].[Close Date]) AS [MaxOfClose Date], SchGrp.SGrpID, Sch"
    "Grp.UnitsInGroup, SchGrp.BLCD, SchGrp.LeaseTermPrimary, DateAdd(\"m\",[LeaseTerm"
    "Primary],[BLCD])-1 AS Expiry, DateDiff(\"d\",[BLCD],(DateAdd(\"m\",[LeaseTermPri"
    "mary],[BLCD]-1))) AS Days, ([LeaseTermPrimary]/12)*[SgprAllow1] AS LTDMI, Units."
    "unitinserv\015\012FROM SchGrp INNER JOIN ([costco repairsNew] INNER JOIN Units O"
    "N [costco repairsNew].Unitref=Units.UnitID) ON SchGrp.SGrpID=Units.GroupREF\015\012"
    "GROUP BY [costco repairsNew].Unitref, SchGrp.SGrpID, SchGrp.UnitsInGroup, SchGrp"
    ".BLCD, SchGrp.LeaseTermPrimary, DateAdd(\"m\",[LeaseTermPrimary],[BLCD])-1, Date"
    "Diff(\"d\",[BLCD],(DateAdd(\"m\",[LeaseTermPrimary],[BLCD]-1))), ([LeaseTermPrim"
    "ary]/12)*[SgprAllow1], Units.unitinserv\015\012HAVING (((Max([costco repairsNew]"
    ".HubMI))>0))\015\012ORDER BY [costco repairsNew].Unitref;\015\012"
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
        dbText "Name" ="MaxOfHubMI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfClose DAte"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[costco repairsNew].Unitref"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbInteger "ColumnWidth" ="1050"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpNoUnits"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrpblcd"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.sgrppterm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expiry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Days"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LTDMI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.unitinserv"
        dbLong "AggregateType" ="-1"
    End
End
