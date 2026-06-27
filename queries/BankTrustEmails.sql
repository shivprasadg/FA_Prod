SELECT
  Location_contacts.clientRef AS BankId,
  Location_contacts.contactFullName AS TCName,
  Location_contacts.contactEmail AS TCEmail,
  Location_contacts.contactClassification
FROM
  Location_contacts
GROUP BY
  Location_contacts.clientRef,
  Location_contacts.contactFullName,
  Location_contacts.contactEmail,
  Location_contacts.contactClassification
HAVING
  (
    (
      (Location_contacts.contactEmail) Is Not Null
    )
    AND (
      (
        Location_contacts.contactClassification
      ) Like "*TR*"
    )
  )
ORDER BY
  Location_contacts.clientRef,
  Location_contacts.contactEmail;
