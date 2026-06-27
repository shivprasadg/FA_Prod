dbMemo "SQL" ="SELECT [clientcompanyName] & IIf([clientDBA]<>\"\",\" d/b/a \" & [clientDBA],\"\""
    ") AS ClientName, *\015\012FROM clients\015\012WHERE (((clients.clientID)<>4133) "
    "AND ((clients.clientType)=\"Remarketing\" Or (clients.clientType) Like \"*RMKT*\""
    "));\015\012"
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
        dbText "Name" ="ClientName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientDBA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientAddress1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientAddress2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientZip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientMainPhone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientFax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientWebsite"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientAnnualSales"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientTaxID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.creditType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.creditMoodyRating"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.StockSym"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1012"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.LeaseSTofCorp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientCompanyName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientBecameDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.[WH?]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.[Dlr/Resale]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.LOinitial"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientOBC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.FleetReplacementCycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.creditFinance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.creditMaint"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.ClientStatus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.PortfolioNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.ClientGroupId"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.Longitude"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.FAVenorNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.ClientCONo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.cpadded"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.SFAccountID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.NMVERIFIED"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.SFDCid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.upsize_ts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.ForceList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.LeaseEntity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientInitContact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.SalespersonREF"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientCoNm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.CLIENTSUB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.CLIENTGRPCNTRY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientcounty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientCountry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.clientIndustry"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.CO-OP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.ClientComm%"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.ONPortal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.WH?"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.CPMModel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.Dlr/Resale"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.ClientCSA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.PIC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.OrigBDEX"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.Latitude"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="clients.AssigneeBank"
        dbLong "AggregateType" ="-1"
    End
End
