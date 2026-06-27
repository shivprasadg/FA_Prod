SELECT
  Contacts.*
FROM
  Contacts
WHERE
  (
    (
      (Contacts.clientRef)= [forms]![bank_view]![clientid]
    )
  );
