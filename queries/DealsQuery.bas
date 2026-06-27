Operation =1
Option =0
Begin InputTables
    Name ="Clients"
    Name ="Deals"
    Name ="lookup_transTypes"
End
Begin OutputColumns
    Alias ="edit"
    Expression ="\"Edit\""
    Expression ="Deals.dealID"
    Expression ="lookup_transTypes.TransType"
    Expression ="Deals.clientRef"
    Expression ="Deals.FAID"
    Alias ="companyAndDBA"
    Expression ="[clientcompanyName] & IIf([clientDBA]<>\"\",\" d/b/a \" & [clientDBA],\"\")"
    Expression ="Deals.Increment"
    Expression ="Deals.dealShortDesc"
    Expression ="Deals.dealRFPDate"
    Expression ="Deals.dealLPAFDue"
    Expression ="Deals.dealDelDate"
    Expression ="Deals.dealDelThru"
    Expression ="Deals.CRfinancialsIH"
    Expression ="Deals.CRFinancialLink"
    Expression ="Deals.CRGuarantor"
    Expression ="Deals.compBank"
    Expression ="Deals.dealLPAFSent"
    Expression ="Deals.dealOutcomeDate"
    Expression ="Deals.dealOutcome"
    Expression ="Deals.dealOutcomeReason"
    Expression ="Deals.dealUnitQty"
    Expression ="Deals.dealUnitPrice"
    Expression ="Deals.dealGrandTotal"
    Expression ="Deals.dealSchedNum"
    Expression ="Deals.transTypeRef"
    Expression ="Deals.dealEnteredDate"
    Expression ="Deals.dealEstDelDate"
    Expression ="Deals.dealMPY"
    Expression ="Deals.dealTerm"
    Expression ="Deals.dealGrandTotal"
    Expression ="Deals.dealLPAFcomp"
    Expression ="Deals.LostTo"
    Expression ="Deals.LostRate"
    Expression ="Deals.LostTerm"
    Expression ="Deals.LostResid"
    Expression ="Deals.[LostImplicit Rt]"
    Expression ="Deals.LostTo"
    Expression ="Deals.DealLOEx"
End
Begin Joins
    LeftTable ="Deals"
    RightTable ="lookup_transTypes"
    Expression ="Deals.transTypeRef = lookup_transTypes.TransTypeID"
    Flag =1
    LeftTable ="Clients"
    RightTable ="Deals"
    Expression ="Clients.clientID = Deals.clientRef"
    Flag =1
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
        dbText "Name" ="edit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.Increment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.FAID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.clientRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealShortDesc"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="5625"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Deals.dealRFPDate"
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
        dbText "Name" ="Deals.dealGrandTotal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.CRfinancialsIH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealSchedNum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealUnitPrice"
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
        dbText "Name" ="Deals.CRFinancialLink"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.CRGuarantor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.compBank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealUnitQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="companyAndDBA"
        dbInteger "ColumnWidth" ="3885"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.transTypeRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealEnteredDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealEstDelDate"
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
        dbText "Name" ="Deals.dealLPAFDue"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFSent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="lookup_transTypes.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1022"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.dealLPAFcomp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.LostTo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.LostRate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.LostTerm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.LostResid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.[LostImplicit Rt]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Deals.DealLOEx"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1258
    Bottom =686
    Left =-1
    Top =-1
    Right =1226
    Bottom =177
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =248
        Top =12
        Right =392
        Bottom =156
        Top =0
        Name ="Clients"
        Name =""
    End
    Begin
        Left =36
        Top =0
        Right =198
        Bottom =214
        Top =0
        Name ="Deals"
        Name =""
    End
    Begin
        Left =440
        Top =12
        Right =584
        Bottom =156
        Top =0
        Name ="lookup_transTypes"
        Name =""
    End
End
