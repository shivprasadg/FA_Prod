SELECT
  DISTINCT TempSC.clientID,
  Sum(1) AS TRANSMGRVISIT
FROM
  Contacts
  INNER JOIN (
    TempSC
    INNER JOIN Notes ON TempSC.clientID = Notes.clientRef
  ) ON Contacts.contactFullName = Notes.ContactNM
WHERE
  (
    (
      (Contacts.contactRTitle)= "4"
    )
    AND (
      (Notes.noteDateTime) Is Not Null
      And (Notes.noteDateTime)> DateAdd(
        "d",
        Date(),
        -365
      )
    )
    AND (
      (Notes.calltype)= "webinar"
      Or (Notes.calltype)= "on-site"
    )
  )
GROUP BY
  TempSC.clientID
ORDER BY
  TempSC.clientID;
