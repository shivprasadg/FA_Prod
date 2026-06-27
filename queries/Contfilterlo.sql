SELECT
  Contacts.*
FROM
  Contacts
WHERE
  (
    (
      (Contacts.clientRef)= [forms]![ClientsViewlo]![clientid]
    )
  );
