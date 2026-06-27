SELECT
  Contacts.contactID,
  Contacts.contactFullName,
  Contacts.contactLast,
  Contacts.contactFirst,
  Contacts.contactTitle,
  Contacts.clientRef
FROM
  Contacts
WHERE
  (
    (
      (Contacts.clientRef)= [Forms]![Deals_view]![clientRef]
    )
  )
ORDER BY
  Contacts.contactLast,
  Contacts.contactFirst;
