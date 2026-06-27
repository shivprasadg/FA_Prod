SELECT
  Contacts.*
FROM
  Contacts
WHERE
  (
    (
      (Contacts.clientRef)= [forms]![ClientsView]![clientid]
    )
  );
