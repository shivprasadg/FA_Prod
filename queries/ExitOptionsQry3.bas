dbMemo "SQL" ="SELECT DISTINCTROW DLookUp(\"UnitUnitNum\",\"Units\",\"UnitId=\" & [EXITOptions_"
    "WorkTable].[UnitId] & \"\") AS UnitNum, DLookUp(\"UnitVin\",\"Units\",\"UnitId=\""
    " & [EXITOptions_WorkTable].[UnitId] & \"\") AS VIN, DLookUp(\"BLCD\",\"SchGrp\","
    "\"SGrpID=\" & [GroupId] & \"\") AS BLCD, DLookUp(\"LeaseTermPrimary\",\"SchGrp\""
    ",\"SGrpID=\" & [GroupId] & \"\") AS LeaseTerm, DateAdd(\"m\",[LeaseTerm],[BLCD])"
    "-1 AS Termination, DateAdd(\"m\",Nz(DLookUp(\"ExchangeITMonth\",\"SchGrp\",\"SGr"
    "pID=\" & [GroupId] & \"\"),0),[BLCD])-1 AS ExchangeITDate, ([Termination]-91) AS"
    " NoticeDate, IIf([ExchangeITDate]=\"TBD\",100,DateDiff(\"m\",Date(),[ExchangeITD"
    "ate])) AS MonthsUnitEXIT, UnitExitOption.GroupId AS Expr1, UnitExitOption.Option"
    "Selected AS Expr2, UnitExitOption.ExtensionStart AS Expr3, UnitExitOption.Extens"
    "ionTerm AS Expr4, UnitExitOption.PrevTermination AS Expr5, UnitExitOption.CurrTe"
    "rmination AS Expr6, UnitExitOption.ExchangeDate AS Expr7, UnitExitOption.Exchang"
    "edMatched AS Expr8, UnitExitOption.ReplaceScheduleId AS Expr9, UnitExitOption.Re"
    "placedByUnitId AS Expr10, UnitExitOption.ExchangePool AS Expr11, UnitExitOption."
    "DateSeleted AS Expr12, UnitExitOption.PrevBLCD AS Expr13, UnitExitOption.LFR AS "
    "Expr14, UnitExitOption.Rent AS Expr15, UnitExitOption.ClientNoticeSent AS Expr16"
    ", UnitExitOption.BankNoticeSent AS Expr17, UnitExitOption.NoticeConfrimed AS Exp"
    "r18, UnitExitOption.AssigneeBankId, UnitExitOption.EstTurnInDate AS Expr19, Unit"
    "ExitOption.ActualTurnInDate AS Expr20, UnitExitOption.OptionStatus AS Expr21\015"
    "\012FROM EXITOptions_WorkTable LEFT JOIN UnitExitOption ON EXITOptions_WorkTable"
    ".UnitID = UnitExitOption.UnitId;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="NoticeDate"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BLCD"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.CurrTermination"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplacedByUnitId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.DateSeleted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.LFR"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionSelected"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ExchangePool"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.GroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.PrevBLCD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ClientNoticeSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ExchangeITDate"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitNum"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ExchangedMatched"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.PrevTermination"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ReplaceScheduleId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.OptionStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ExchangeDate"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ExtensionTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ExtensionStart"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.BankNoticeSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.EstTurnInDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MonthsUnitEXIT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.ActualTurnInDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.Rent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LeaseTerm"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.NoticeConfrimed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VIN"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Termination"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitExitOption.AssigneeBankId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr21"
        dbLong "AggregateType" ="-1"
    End
End
