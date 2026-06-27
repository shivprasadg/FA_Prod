SELECT
  Contacts.*
FROM
  Contacts
WHERE
  (
    (
      (Contacts.clientRef)= [forms]![RemarketingClientView]![clientid]
    )
  );
