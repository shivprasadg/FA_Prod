dbMemo "SQL" ="SELECT Min(BankCustEmails0.CustName) AS CustName, Min(BankCustEmails0.CustEmail)"
    " AS CustEmail, Min(BankCustEmails0.CustCName) AS CustCName, Min(BankCustEmails0."
    "CustCEmail) AS CustCEmail, Min(BankCustEmails0.ABSEmail) AS ABSEmail\015\012FROM"
    " (SELECT Location_contacts.clientRef AS BankId, IIf([contactClassification]=\"Cu"
    "stodian\",[contactFullName],Null) AS CustName, IIf([contactClassification]=\"Cus"
    "todian\",[contactEmail],Null) AS CustEmail, IIf([contactClassification]=\"Custod"
    "ian-Contact\",[contactFullName],Null) AS CustCName, IIf([contactClassification]="
    "\"Custodian-Contact\",[contactEmail],Null) AS CustCEmail, IIf([contactClassifica"
    "tion]=\"Custodian-Vault\",[contactEmail],Null) AS ABSEmail FROM Location_contact"
    "s WHERE (((Location_contacts.clientRef)=5233) AND ((Location_contacts.contactCla"
    "ssification) Like \"*Custodian*\")) ORDER BY Location_contacts.clientRef, IIf([c"
    "ontactClassification]=\"Custodian\",[contactEmail],Null))  AS BankCustEmails0;\015"
    "\012"
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
        dbText "Name" ="CustEmail"
        dbInteger "ColumnWidth" ="3765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustCName"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustName"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CustCEmail"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ABSEmail"
        dbInteger "ColumnWidth" ="3420"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
