SELECT
  Contacts.contactID,
  Contacts.clientRef
FROM
  Contacts
WHERE
  (
    (
      (Contacts.contactID) In (
        SELECT
          [contactID]
        FROM
          [Contacts] As Tmp
        GROUP BY
          [contactID]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  Contacts.contactID;
