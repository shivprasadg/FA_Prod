dbMemo "SQL" ="Select ClientID, MLID, SignerGroup1 As SignerType, SchID, SGrpID, MLSignerID As "
    "ContactID FROM ATF_FAID_ContactsMLSigners WHERE SignerGroup1='Master Lease'\015\012"
    "Union ALL\015\012Select ClientID, MLID, SignerGroup1 As SignerType, SchID, SGrpI"
    "D, MLSigner2ID As ContactID FROM ATF_FAID_ContactsMLSigners WHERE SignerGroup1='"
    "Master Lease'\015\012Union ALL\015\012Select ClientID, MLID, SignerGroup1 As Sig"
    "nerType, SchID, SGrpID, MLSigner3ID As ContactID FROM ATF_FAID_ContactsMLSigners"
    " WHERE SignerGroup1='Master Lease'\015\012Union ALL\015\012Select ClientID, MLID"
    ", SignerGroup2 As SignerType, SchID, SGrpID, MLGuarSignerID As ContactID FROM AT"
    "F_FAID_ContactsMLSigners WHERE SignerGroup2='Guarantor';\015\012Union ALL\015\012"
    "Select ClientID, MLID, SignerGroup3 As SignerType, SchID, SGrpID, SchSignerID As"
    " ContactID FROM ATF_FAID_ContactsMLSigners WHERE SignerGroup3='Schedule';\015\012"
    "Union ALL\015\012Select ClientID, MLID, SignerGroup3 As SignerType, SchID, SGrpI"
    "D, SchSigner2ID As ContactID FROM ATF_FAID_ContactsMLSigners WHERE SignerGroup3="
    "'Schedule';\015\012Union ALL\015\012Select ClientID, MLID, SignerGroup3 As Signe"
    "rType, SchID, SGrpID, SchSigner3ID As ContactID FROM ATF_FAID_ContactsMLSigners "
    "WHERE SignerGroup3='Schedule';\015\012UNION ALL Select ClientID, MLID, SignerGro"
    "up4 As SignerType, SchID, SGrpID, CofASignerID As ContactID FROM ATF_FAID_Contac"
    "tsMLSigners WHERE SignerGroup4='COA';\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="SignerType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClientID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SchID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SGrpID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ContactID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MLID"
        dbLong "AggregateType" ="-1"
    End
End
