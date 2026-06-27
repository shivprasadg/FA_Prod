dbMemo "SQL" ="SELECT Units.GroupREF, Units.UnitID, Units.Unitinserv, MaxMiQ.MaxOfHubMI, MaxMiQ"
    ".MaxOfMiClDt, IIf(IsNull([MaxOfReefer Hours]),0,[maxofreefer hours]) AS MaxHrs, "
    "MaxHrsQ.[MaxOfClose DAte] AS MaxHrsDt, DateAdd(\"m\",[LeaseTermPrimary],[BLCD])-"
    "1 AS Expiry, DateDiff(\"d\",[unitinserv],[MaxofMIclDt])+1 AS MiDays, IIf(IsNull("
    "[maxofclose date]),0,DateDiff(\"d\",[unitinserv],[maxofclose date])+1) AS HrDays"
    ", [MaxOfHubMI]/(DateDiff(\"d\",[unitinserv],[MaxofMIclDt])+1) AS Dailymi, IIf(Is"
    "Null([maxofreefer hours]),0,[maxofreefer hours]/(DateDiff(\"d\",[unitinserv],[ma"
    "xofclose date])+1)) AS DailyHr, DateDiff(\"d\",[unitinserv],(DateAdd(\"m\",[Leas"
    "eTermPrimary],[BLCD])-1)) AS Leasedays, [sgprallow1]*[LeaseTermPrimary]/12 AS Mi"
    "Allowance, (DateDiff(\"d\",[unitinserv],(DateAdd(\"m\",[LeaseTermPrimary],[BLCD]"
    ")-1)))*([MaxOfHubMI]/(DateDiff(\"d\",[unitinserv],[MaxofMIclDt])+1)) AS ProjLTMI"
    ", [sgprallow4]*[LeaseTermPrimary]/12 AS HrAllowance, (DateDiff(\"d\",[unitinserv"
    "],(DateAdd(\"m\",[LeaseTermPrimary],[BLCD])-1)))*(IIf(IsNull([maxofreefer hours]"
    "),0,[maxofreefer hours]/(DateDiff(\"d\",[unitinserv],[maxofclose date])+1))) AS "
    "ProjLTDHr, Units.UnitUnitNum, ((DateDiff(\"d\",[unitinserv],(DateAdd(\"m\",[Leas"
    "eTermPrimary],[BLCD])-1)))*([MaxOfHubMI]/(DateDiff(\"d\",[unitinserv],[MaxofMIcl"
    "Dt])+1)))/[LeaseTermPrimary] AS ProjMoMi, ((DateDiff(\"d\",[unitinserv],(DateAdd"
    "(\"m\",[LeaseTermPrimary],[BLCD])-1)))*(IIf(IsNull([maxofreefer hours]),0,[maxof"
    "reefer hours]/(DateDiff(\"d\",[unitinserv],[maxofclose date])+1))))/[LeaseTermPr"
    "imary] AS ProjMoHr, [SGPRALLOW1]/12 AS LeaseMoMI, [sgPRallow4]/12 AS LeaseMoHr\015"
    "\012FROM ((MstrLease INNER JOIN Schedule ON MstrLease.MLID=Schedule.SchMLRef) IN"
    "NER JOIN SchGrp ON Schedule.SchID=SchGrp.SGrpSchID) INNER JOIN ((Units LEFT JOIN"
    " MaxMiQ ON Units.UnitID=MaxMiQ.Unitref) LEFT JOIN MaxHrsQ ON Units.UnitID=MaxHrs"
    "Q.Unitref) ON SchGrp.SGrpID=Units.GroupREF\015\012WHERE (((Schedule.SchNo) Like "
    "\"*toll*\") AND ((MstrLease.MLNo) Like \"Costco*\"))\015\012ORDER BY Schedule.Sc"
    "hNo, Units.GroupREF, Units.UnitID;\015\012"
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
        dbText "Name" ="MaxHrs"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="Fixed"
        dbText "InputMask" ="\\2"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Units.UnitID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="Units.GroupREF"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1320"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="MaxMiQ.MaxOfHubMI"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="2"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="MaxMiQ.MaxOfMiClDt"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="mm/dd/yy"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="MaxHrsDt"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1320"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="mm/dd/yy"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="Expiry"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="mm/dd/yy"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="MiDays"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="Units.Unitinserv"
        dbLong "AggregateType" ="-1"
        dbText "Format" ="mm/dd/yy"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="HrDays"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="ProjLTDHr"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
    End
    Begin
        dbText "Name" ="Dailymi"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="DailyHr"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Leasedays"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="ProjLTMI"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="MiAllowance"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
    End
    Begin
        dbText "Name" ="HrAllowance"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Units.UnitUnitNum"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjMoMi"
        dbInteger "ColumnWidth" ="1815"
        dbInteger "ColumnOrder" ="19"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ProjMoHr"
        dbInteger "ColumnOrder" ="21"
        dbText "Format" ="Fixed"
        dbByte "DecimalPlaces" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseMoMI"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseMoHr"
        dbLong "AggregateType" ="-1"
    End
End
