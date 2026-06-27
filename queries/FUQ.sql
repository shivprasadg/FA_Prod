SELECT
  Clients.ClientShNm,
  Salesperson.SPInitials,
  Notes.noteDateTime,
  Notes.noteFU,
  Notes.noteMemo,
  Contacts.contactFullName
FROM
  (
    (
      Notes
      INNER JOIN Clients ON Notes.clientRef = Clients.clientID
    )
    INNER JOIN Salesperson ON Clients.SalespersonREF = Salesperson.SalespersonID
  )
  LEFT JOIN Contacts ON Notes.MAINCONTACT = Contacts.contactID
WHERE
  (
    (
      (Salesperson.SPInitials)= forms!funotef!fuloex
    )
    And (
      (Notes.noteFU) Is Not Null
      And (Notes.noteFU)<= forms!funotef!fupdt
    )
  );
