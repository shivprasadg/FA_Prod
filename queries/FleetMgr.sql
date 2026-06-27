SELECT
  TempSC.clientID,
  Sum(1) AS FMgrvisit
FROM
  (
    TempSC
    INNER JOIN Notes ON TempSC.clientID = Notes.clientRef
  )
  INNER JOIN Contacts ON Notes.ContactNM = Contacts.contactFullName
WHERE
  (
    (
      (Notes.noteDateTime) Is Not Null
      And (Notes.noteDateTime)> DateAdd(
        "d",
        Date(),
        -365
      )
    )
    AND (
      (Notes.CallType)= "on-site"
      Or (Notes.CallType)= "webinar"
    )
    AND (
      (Contacts.contactRTitle)= "5"
    )
  )
GROUP BY
  TempSC.clientID;
