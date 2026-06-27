Operation =1
Option =0
Having ="(((SchGrp.Assigned_Indicator)=\"F\" Or (SchGrp.Assigned_Indicator)=\"P\") AND (("
    "SchGrp.SynFee)=0) AND ((vw_SixKeys.MLOrig)=\"FA\")) OR (((SchGrp.Assigned_Indica"
    "tor)=\"F\" Or (SchGrp.Assigned_Indicator)=\"P\") AND ((SchGrp.Fee_Projected)=0) "
    "AND ((vw_SixKeys.MLOrig)=\"FA\")) OR (((SchGrp.Assigned_Indicator)=\"F\" Or (Sch"
    "Grp.Assigned_Indicator)=\"P\") AND ((SchGrp.SGrpSalesIndex) Is Null) AND ((vw_Si"
    "xKeys.MLOrig)=\"FA\")) OR (((SchGrp.Assigned_Indicator)=\"F\" Or (SchGrp.Assigne"
    "d_Indicator)=\"P\") AND ((SchGrp.SGrpIndex) Is Null) AND ((vw_SixKeys.MLOrig)=\""
    "FA\")) OR (((SchGrp.Assigned_Indicator)=\"F\" Or (SchGrp.Assigned_Indicator)=\"P"
    "\") AND ((SchGrp.SgrpYield) Is Null) AND ((vw_SixKeys.MLOrig)=\"FA\")) OR (((Sch"
    "Grp.Assigned_Indicator)=\"F\" Or (SchGrp.Assigned_Indicator)=\"P\") AND ((SchGrp"
    ".SgrpSalesYield) Is Null) AND ((vw_SixKeys.MLOrig)=\"FA\")) OR (((SchGrp.Assigne"
    "d_Indicator)=\"F\" Or (SchGrp.Assigned_Indicator)=\"P\") AND ((SchGrp.SGrpRALB) "
    "Is Null) AND ((vw_SixKeys.MLOrig)=\"FA\")) OR (((SchGrp.Assigned_Indicator)=\"F\""
    " Or (SchGrp.Assigned_Indicator)=\"P\") AND ((Schedule.SchIn) Is Null) AND ((vw_S"
    "ixKeys.MLOrig)=\"FA\")) OR (((SchGrp.Assigned_Indicator)=\"F\" Or (SchGrp.Assign"
    "ed_Indicator)=\"P\") AND ((IIf([SchGrp].[SGRPNoPD]=-1 And Sum(Nz([UnitPDRent],0)"
    ")=0,DLookUp(\"Person\",\"GroupFinalizedPersons\",\"Fieldname='PDRent'\"))) Is No"
    "t Null) AND ((vw_SixKeys.MLOrig)=\"FA\")) OR (((SchGrp.Assigned_Indicator)=\"F\""
    " Or (SchGrp.Assigned_Indicator)=\"P\") AND ((IIf(Nz([SchGrp].[BillToAddressID],\""
    "\")=\"\",DLookUp(\"Person\",\"GroupFinalizedPersons\",\"Fieldname='BillToAddress"
    "ID'\"))) Is Not Null) AND ((vw_SixKeys.MLOrig)=\"FA\")) OR (((SchGrp.Assigned_In"
    "dicator)=\"F\" Or (SchGrp.Assigned_Indicator)=\"P\") AND ((IIf(Nz([SchGrp].[Sgrp"
    "CommInd],\"\")=\"\",DLookUp(\"Person\",\"GroupFinalizedPersons\",\"Fieldname='Sg"
    "rpCommInd'\"))) Is Not Null) AND ((vw_SixKeys.MLOrig)=\"FA\")) OR (((SchGrp.Assi"
    "gned_Indicator)=\"F\" Or (SchGrp.Assigned_Indicator)=\"P\") AND ((IIf(Sum(Nz([Un"
    "itRent],0))=0,DLookUp(\"Person\",\"GroupFinalizedPersons\",\"Fieldname='UnitRent"
    "'\"))) Is Not Null) AND ((vw_SixKeys.MLOrig)=\"FA\")) OR (((SchGrp.Assigned_Indi"
    "cator)=\"F\" Or (SchGrp.Assigned_Indicator)=\"P\") AND ((IIf(Sum(Nz([UnitSynResi"
    "d],0))=0,DLookUp(\"Person\",\"GroupFinalizedPersons\",\"Fieldname='UnitSynResid'"
    "\"))) Is Not Null) AND ((vw_SixKeys.MLOrig)=\"FA\"))"
Begin InputTables
    Name ="SchGrp"
    Name ="vw_SixKeys"
    Name ="Schedule"
    Name ="Units"
End
Begin OutputColumns
    Alias ="Client"
    Expression ="vw_SixKeys.ClientShortName"
    Expression ="vw_SixKeys.Schedule"
    Expression ="vw_SixKeys.UnitGroup"
    Alias ="Units"
    Expression ="Count(vw_SixKeys.UnitId)"
    Expression ="SchGrp.AssignmentDate"
    Expression ="SchGrp.Assigned_Indicator"
    Expression ="SchGrp.SynFee"
    Expression ="SchGrp.Fee_Projected"
    Expression ="SchGrp.SGrpSalesIndex"
    Expression ="SchGrp.SGrpIndex"
    Expression ="SchGrp.SgrpYield"
    Expression ="SchGrp.SgrpSalesYield"
    Expression ="SchGrp.SGrpRALB"
    Expression ="Schedule.SchIn"
    Alias ="PDiemRent"
    Expression ="IIf([SchGrp].[SGRPNoPD]=-1 And Sum(Nz([UnitPDRent],0))=0,DLookUp(\"Person\",\"Gr"
        "oupFinalizedPersons\",\"Fieldname='PDRent'\"))"
    Alias ="BillToAddress"
    Expression ="IIf(Nz([SchGrp].[BillToAddressID],\"\")=\"\",DLookUp(\"Person\",\"GroupFinalized"
        "Persons\",\"Fieldname='BillToAddressID'\"))"
    Alias ="BLCDBasis"
    Expression ="IIf(Nz([SchGrp].[SgrpCommInd],\"\")=\"\",DLookUp(\"Person\",\"GroupFinalizedPers"
        "ons\",\"Fieldname='SgrpCommInd'\"))"
    Alias ="UnitSyndRent"
    Expression ="IIf(Sum(Nz([UnitRent],0))=0,DLookUp(\"Person\",\"GroupFinalizedPersons\",\"Field"
        "name='UnitRent'\"))"
    Alias ="UnitResidual"
    Expression ="IIf(Sum(Nz([UnitSynResid],0))=0,DLookUp(\"Person\",\"GroupFinalizedPersons\",\"F"
        "ieldname='UnitSynResid'\"))"
    Expression ="Schedule.SchID"
    Expression ="SchGrp.SGrpID"
End
Begin Joins
    LeftTable ="SchGrp"
    RightTable ="vw_SixKeys"
    Expression ="SchGrp.SGrpID = vw_SixKeys.SGrpID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Schedule"
    Expression ="vw_SixKeys.SchID = Schedule.SchID"
    Flag =1
    LeftTable ="vw_SixKeys"
    RightTable ="Units"
    Expression ="vw_SixKeys.UnitID = Units.UnitID"
    Flag =1
End
Begin OrderBy
    Expression ="vw_SixKeys.ClientShortName"
    Flag =0
    Expression ="vw_SixKeys.Schedule"
    Flag =0
    Expression ="vw_SixKeys.UnitGroup"
    Flag =0
End
Begin Groups
    Expression ="vw_SixKeys.ClientShortName"
    GroupLevel =0
    Expression ="vw_SixKeys.Schedule"
    GroupLevel =0
    Expression ="vw_SixKeys.UnitGroup"
    GroupLevel =0
    Expression ="SchGrp.AssignmentDate"
    GroupLevel =0
    Expression ="SchGrp.Assigned_Indicator"
    GroupLevel =0
    Expression ="SchGrp.SynFee"
    GroupLevel =0
    Expression ="SchGrp.Fee_Projected"
    GroupLevel =0
    Expression ="SchGrp.SGrpSalesIndex"
    GroupLevel =0
    Expression ="SchGrp.SGrpIndex"
    GroupLevel =0
    Expression ="SchGrp.SgrpYield"
    GroupLevel =0
    Expression ="SchGrp.SgrpSalesYield"
    GroupLevel =0
    Expression ="SchGrp.SGrpRALB"
    GroupLevel =0
    Expression ="Schedule.SchIn"
    GroupLevel =0
    Expression ="IIf(Nz([SchGrp].[BillToAddressID],\"\")=\"\",DLookUp(\"Person\",\"GroupFinalized"
        "Persons\",\"Fieldname='BillToAddressID'\"))"
    GroupLevel =0
    Expression ="IIf(Nz([SchGrp].[SgrpCommInd],\"\")=\"\",DLookUp(\"Person\",\"GroupFinalizedPers"
        "ons\",\"Fieldname='SgrpCommInd'\"))"
    GroupLevel =0
    Expression ="vw_SixKeys.MLOrig"
    GroupLevel =0
    Expression ="SchGrp.SGRPNoPD"
    GroupLevel =0
    Expression ="Schedule.SchID"
    GroupLevel =0
    Expression ="SchGrp.SGrpID"
    GroupLevel =0
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
        dbText "Name" ="vw_SixKeys.Schedule"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="vw_SixKeys.UnitGroup"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="SchGrp.SynFee"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="Client"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="Units"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="SchGrp.Assigned_Indicator"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="Schedule.SchIn"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpSalesYield"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
    End
    Begin
        dbText "Name" ="SchGrp.Fee_Projected"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpSalesIndex"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpRALB"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="SchGrp.SgrpYield"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpIndex"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="PDiemRent"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="BLCDBasis"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="BillToAddress"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitResidual"
        dbInteger "ColumnWidth" ="1605"
        dbInteger "ColumnOrder" ="18"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UnitSyndRent"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Units.UNITPDRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfUNITPDRENT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGRPNoPD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.AssignmentDate"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Schedule.SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchGrp.SGrpID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =83
    Top =39
    Right =1407
    Bottom =968
    Left =-1
    Top =-1
    Right =1306
    Bottom =230
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =241
        Top =91
        Right =385
        Bottom =555
        Top =0
        Name ="SchGrp"
        Name =""
    End
    Begin
        Left =59
        Top =14
        Right =203
        Bottom =335
        Top =0
        Name ="vw_SixKeys"
        Name =""
    End
    Begin
        Left =462
        Top =13
        Right =606
        Bottom =421
        Top =0
        Name ="Schedule"
        Name =""
    End
    Begin
        Left =654
        Top =12
        Right =798
        Bottom =233
        Top =0
        Name ="Units"
        Name =""
    End
End
