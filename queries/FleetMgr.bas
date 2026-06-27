dbMemo "SQL" ="SELECT TempSC.clientID, Sum(1) AS FMgrvisit\015\012FROM (TempSC INNER JOIN Notes"
    " ON TempSC.clientID=Notes.clientRef) INNER JOIN Contacts ON Notes.ContactNM=Cont"
    "acts.contactFullName\015\012WHERE (((Notes.noteDateTime) Is Not Null And (Notes."
    "noteDateTime)>DateAdd(\"d\",Date(),-365)) AND ((Notes.CallType)=\"on-site\" Or ("
    "Notes.CallType)=\"webinar\") AND ((Contacts.contactRTitle)=\"5\"))\015\012GROUP "
    "BY TempSC.clientID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="FMgrvisit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TempSC.clientID"
        dbLong "AggregateType" ="-1"
    End
End
