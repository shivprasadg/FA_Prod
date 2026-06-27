SELECT
  client_lookup.CompanyAndDBA,
  Contacts.contactClassification,
  Contacts.contactFullName,
  Contacts.contactEmail,
  Contacts.contactPrimaryPhone,
  Contacts.contactPhoneExt,
  Contacts.contactCell,
  Contacts.contactPhone2,
  Contacts.contactPrefix,
  Contacts.contactFirst,
  Contacts.contactMiddle,
  Contacts.contactLast,
  Contacts.contactSuffix,
  Contacts.contactTitle,
  Contacts.contactOptOut,
  Clients.clientType
FROM
  (
    Contacts
    INNER JOIN client_lookup ON Contacts.clientRef = client_lookup.clientID
  )
  INNER JOIN Clients ON client_lookup.clientID = Clients.clientID
WHERE
  (
    (
      (Clients.clientType)<> "dead"
    )
    AND (
      (Contacts.ContactObsolete)= No
    )
  );
