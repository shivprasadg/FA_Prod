Operation =1
Option =0
Where ="(((Salesperson.SalespersonName) Like \"s*\" Or (Salesperson.SalespersonName) Lik"
    "e \"H*\") AND ((Deals.transTypeRef)<>1) AND ((Deals.dealOutcome) Is Not Null And"
    " ((Deals.dealOutcome)=\"dead\" Or (Deals.dealOutcome)=\"lost\")) AND ((Deals.dea"
    "lOutcomeDate)>=#1/1/2012# And (Deals.dealOutcomeDate)<=#3/31/2012#) AND ((Deals."
    "clientRef)<>4133) AND ((lookup_transTypes.TransType)=\"Proposal\") AND ((Groups."
    "AltEquip)=No))"
Begin InputTables
    Name ="Clients"
    Name ="Salesperson"
    Name ="Deals"
    Name ="lookup_transTypes"
    Name ="lookup_GroupMakes"
    Name ="Groups"
    Name ="lookup_groupModels"
    Name ="lookup_groupTypes"
End
Begin OutputColumns
    Expression ="Salesperson.SalespersonName"
    Expression ="Deals.transTypeRef"
    Expression ="Deals.dealOutcome"
    Expression ="Deals.dealOutcomeDate"
    Expression ="Deals.dealLPAFSent"
    Expression ="Deals.dealID"
    Expression ="Deals.clientRef"
    Expression ="Deals.dealRFPDate"
    Expression ="Deals.dealLPAFSent"
    Expression ="Clients.clientCompanyName"
    Expression ="Deals.dealUnitQty"
    Expression ="Groups.groupYear"
    Expression ="Groups.groupMake"
    Expression ="Groups.groupModel"
    Alias ="Type"
    Expression ="IIf(Left([groups].[groupType],6)=\"daycab\",\"Daycab\",[groups].[grouptype])"
    Expression ="Deals.dealGrandTotal"
    Expression ="Groups.groupTotal"
    Expression ="Deals.dealSchedNum"
    Expression ="Deals.dealOutcomeReason"
    Expression ="Groups.groupUnits"
    Expression ="lookup_transTypes.TransType"
    Alias ="UQTY"
    Expression ="IIf(IsNull([groupunits]),[dealunitqty],[groupunits])"
    Alias ="D$"
    Expression ="IIf(IsNull([grouptotal]),[dealgrandtotal],[grouptotal])"
    Expression ="Deals.FAID"
    Expression ="Groups.AltEquip"
    Expression ="Salesperson.SalespersonID"
    Expression ="Clients.clientBecameDate"
    Alias ="Age"
    Expression ="DateDiff(\"d\",[clientbecamedate],Date())"
    Alias ="RptGrp"
    Expression ="IIf(DateDiff(\"d\",[clientbecamedate],Date())<=365,2,1)"
    Expression ="Groups.groupType"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Salesperson"
    Expression ="Clients.SalespersonREF = Salesperson.SalespersonID"
    Flag =1
    LeftTable ="Deals"
    RightTable ="lookup_transTypes"
    Expression ="Deals.transTypeRef = lookup_transTypes.TransTypeID"
    Flag =1
    LeftTable ="Clients"
    RightTable ="Deals"
    Expression ="Clients.clientID = Deals.clientRef"
    Flag =1
    LeftTable ="lookup_GroupMakes"
    RightTable ="Groups"
    Expression ="lookup_GroupMakes.groupMakeID = Groups.groupMakeRef"
    Flag =3
    LeftTable ="Groups"
    RightTable ="lookup_groupModels"
    Expression ="Groups.groupModelRef = lookup_groupModels.groupModelID"
    Flag =2
    LeftTable ="Groups"
    RightTable ="lookup_groupTypes"
    Expression ="Groups.groupTypeRef = lookup_groupTypes.groupTypeID"
    Flag =2
    LeftTable ="Deals"
    RightTable ="Groups"
    Expression ="Deals.dealID = Groups.dealRef"
    Flag =2
End
Begin OrderBy
    Expression ="Salesperson.SalespersonName"
    Flag =0
    Expression ="Deals.dealOutcome"
    Flag =0
    Expression ="Deals.dealID"
    Flag =0
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
        dbText "Name" ="Deals.dealID"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.clientRef"
        dbInteger "ColumnOrder" ="7"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealRFPDate"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeDate"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeReason"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCompanyName"
        dbInteger "ColumnWidth" ="3765"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealUnitQty"
        dbInteger "ColumnOrder" ="19"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupYear"
        dbInteger "ColumnOrder" ="13"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealGrandTotal"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.transTypeRef"
        dbInteger "ColumnOrder" ="11"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFSent"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealSchedNum"
        dbInteger "ColumnOrder" ="14"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupTotal"
        dbInteger "ColumnOrder" ="18"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupUnits"
        dbInteger "ColumnOrder" ="21"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_transTypes.TransType"
        dbInteger "ColumnOrder" ="22"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UQTY"
        dbInteger "ColumnOrder" ="23"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="D$"
        dbInteger "ColumnOrder" ="24"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.FAID"
        dbInteger "ColumnOrder" ="25"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.AltEquip"
        dbInteger "ColumnOrder" ="26"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupMake"
        dbInteger "ColumnOrder" ="15"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupModel"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Salesperson.SalespersonName"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientBecameDate"
        dbInteger "ColumnOrder" ="29"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Salesperson.SalespersonID"
        dbInteger "ColumnOrder" ="28"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Age"
        dbInteger "ColumnOrder" ="20"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RptGrp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Type"
        dbInteger "ColumnOrder" ="17"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Groups.groupType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1004"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1247
    Bottom =679
    Left =-1
    Top =-1
    Right =1215
    Bottom =-1
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =25
        Top =10
        Right =201
        Bottom =245
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =213
        Top =174
        Right =357
        Bottom =318
        Top =0
        Name ="Salesperson"
        Name =""
    End
    Begin
        Left =197
        Top =10
        Right =365
        Bottom =160
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =1027
        Top =12
        Right =1171
        Bottom =156
        Top =0
        Name ="lookup_transTypes"
        Name =""
    End
    Begin
        Left =585
        Top =8
        Right =729
        Bottom =152
        Top =0
        Name ="lookup_GroupMakes"
        Name =""
    End
    Begin
        Left =396
        Top =14
        Right =540
        Bottom =159
        Top =0
        Name ="Groups"
        Name =""
    End
    Begin
        Left =607
        Top =164
        Right =751
        Bottom =308
        Top =0
        Name ="lookup_groupModels"
        Name =""
    End
    Begin
        Left =835
        Top =139
        Right =979
        Bottom =283
        Top =0
        Name ="lookup_groupTypes"
        Name =""
    End
End
