Operation =1
Option =0
Begin InputTables
    Name ="Clients"
    Name ="Deals"
End
Begin OutputColumns
    Expression ="Clients.clientCompanyName"
    Expression ="Deals.transTypeRef"
    Expression ="Deals.dealShortDesc"
    Expression ="Deals.dealRFPDate"
    Alias ="Expr1"
    Expression ="Deals.dealRFPDue"
    Alias ="Expr2"
    Expression ="Deals.dealRFPSent"
    Expression ="Deals.dealOutcomeDate"
    Expression ="Deals.dealOutcome"
    Expression ="Deals.dealOutcomeReason"
    Expression ="Deals.dealUnitQty"
    Expression ="Deals.dealMPY"
    Expression ="Deals.dealTerm"
    Expression ="Deals.dealUnitPrice"
    Expression ="Deals.dealGrandTotal"
    Expression ="Deals.dealEstDelDate"
    Expression ="Deals.dealDelDate"
    Expression ="Deals.dealDelThru"
    Expression ="Deals.CRfinancialsIH"
    Expression ="Deals.CRFinancialLink"
    Expression ="Deals.CRGuarantor"
    Expression ="Deals.CRGtyDate"
    Expression ="Deals.dealSchedSent"
    Expression ="Deals.dealSchedBack"
    Expression ="Deals.dealRALSent"
    Expression ="Deals.dealRALBack"
    Alias ="Expr3"
    Expression ="Deals.compAmeriquest"
    Alias ="Expr4"
    Expression ="Deals.compLeaseplan"
    Alias ="Expr5"
    Expression ="Deals.compFSL"
    Alias ="Expr6"
    Expression ="Deals.compPHH"
    Expression ="Deals.compBank"
    Alias ="Expr7"
    Expression ="Deals.compOther"
    Expression ="Deals.FAID"
    Expression ="Deals.dealSchedNumOrig"
    Expression ="Deals.dealSchedNum"
    Expression ="Deals.Proposaltimeline"
    Expression ="Deals.pipelineOrder"
End
Begin Joins
    LeftTable ="Clients"
    RightTable ="Deals"
    Expression ="Clients.clientID = Deals.clientRef"
    Flag =1
End
Begin OrderBy
    Expression ="Clients.clientCompanyName"
    Flag =0
    Expression ="Deals.transTypeRef"
    Flag =0
    Expression ="Deals.dealOutcomeDate"
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
        dbText "Name" ="Clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.transTypeRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealShortDesc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.pipelineOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealRFPDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealRFPDue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealRFPSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealOutcomeReason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealUnitQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealMPY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealUnitPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealGrandTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealEstDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealDelDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealDelThru"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.CRfinancialsIH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.CRFinancialLink"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.CRGuarantor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.CRGtyDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealSchedSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealSchedBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealRALSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealRALBack"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.compAmeriquest"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.compLeaseplan"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.compFSL"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.compPHH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.compBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.compOther"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealSchedNumOrig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealSchedNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.proposaltimeline"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1152
    Bottom =678
    Left =-1
    Top =-1
    Right =1120
    Bottom =335
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =493
        Bottom =272
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =230
        Bottom =269
        Top =0
        Name ="Deals"
        Name =""
    End
End
