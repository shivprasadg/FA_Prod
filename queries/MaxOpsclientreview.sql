SELECT
  Clients.clientID,
  Notes.CallType,
  Contacts.contactClassification,
  Max(
    IIf(
      IsNull([noteectdt]),
      [notedatetime],
      [noteectdt]
    )
  ) AS DT
FROM
  (
    Notes
    INNER JOIN Clients ON Notes.clientRef = Clients.clientID
  )
  INNER JOIN Contacts ON Notes.ContactNM = Contacts.contactFullName
GROUP BY
  Clients.clientID,
  Notes.CallType,
  Contacts.contactClassification
HAVING
  (
    (
      (Notes.CallType)= "on-site"
    )
    AND (
      (Contacts.contactClassification)= "Operations"
    )
  )
ORDER BY
  Clients.clientID;
