dbMemo "SQL" ="SELECT DISTINCT TempSC.clientID, Sum(1) AS TRANSMGRVISIT\015\012FROM Contacts IN"
    "NER JOIN (TempSC INNER JOIN Notes ON TempSC.clientID=Notes.clientRef) ON Contact"
    "s.contactFullName=Notes.ContactNM\015\012WHERE (((Contacts.contactRTitle)=\"4\")"
    " AND ((Notes.noteDateTime) Is Not Null And (Notes.noteDateTime)>DateAdd(\"d\",Da"
    "te(),-365)) AND ((Notes.calltype)=\"webinar\" Or (Notes.calltype)=\"on-site\"))\015"
    "\012GROUP BY TempSC.clientID\015\012ORDER BY TempSC.clientID;\015\012"
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
        dbText "Name" ="TRANSMGRVISIT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TempSC.clientID"
        dbLong "AggregateType" ="-1"
    End
End
