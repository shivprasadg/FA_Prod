dbMemo "SQL" ="SELECT Itinerary.FrDt, Itinerary.FrTime, Itinerary.ToTime, Itinerary.TimeZone, I"
    "tinerary.Activity, Itinerary.Comment, IIf(IsNull([city]),IIf(IsNull([LCity]),[cl"
    "ientcity],[LCity]),[city]) & \" \" & IIf(IsNull([state]),IIf(IsNull([locations]."
    "[LState]),[clientstate],[locations].[LState]),[state]) AS CityST, Salesperson.SP"
    "Initials, Itinerary.LOEXRef, \"Others: \" & IIf([ab]=Yes,\"AB \",Null) & IIf([mm"
    "]=Yes,\"MM \",Null) & IIf([sm]=Yes,\"SM \",IIf([pf]=Yes,\"PF \",Null) & IIf([JF]"
    "=Yes,\" JF \",Null) & IIf([du]=Yes,\"DU \",Null)) AS Others, Itinerary.Allday, I"
    "tinerary.ItinID, Itinerary.City, Itinerary.State, Itinerary.Allday, Itinerary.Co"
    "ntact, Itinerary.ClientProsp, Clients.ClientShNm, Clients.clientCity, Clients.cl"
    "ientState, Contacts.contactFullName, Locations.LCity, Locations.LState, Clients."
    "creditFAProfile, Clients.creditFAScore, Clients.creditRank, Contacts.contacttitl"
    "e\015\012FROM (((Salesperson INNER JOIN Itinerary ON Salesperson.SalespersonID=I"
    "tinerary.LOEXRef) LEFT JOIN Contacts ON Itinerary.Contact=Contacts.contactID) LE"
    "FT JOIN Locations ON Contacts.cLocationID=Locations.LocationID) LEFT JOIN Client"
    "s ON Itinerary.ClientProsp=Clients.clientID\015\012WHERE (((Itinerary.FrDt)>=[fo"
    "rms]![itinupdtF]![itinstart] And (Itinerary.FrDt)<=[forms]![itinupdtf]![itinend]"
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
        dbText "Name" ="Itinerary.FrDt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Others"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.FrTime"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.ToTime"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.TimeZone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.Activity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.Comment"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Salesperson.SPInitials"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.LOEXRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CityST"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.Allday"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.ItinID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.State"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.Contact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Itinerary.ClientProsp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1010"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.ClientShNm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.clientState"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contactFullName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.LocationCity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Locations.stateRef"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFAProfile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditFAScore"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Clients.creditRank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Contacts.contacttitle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
End
