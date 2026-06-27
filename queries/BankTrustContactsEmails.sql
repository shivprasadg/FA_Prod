SELECT
  LC.clientRef AS BankId,
  LC.contactFullName AS TrustContactName,
  LC.contactEmail AS TrustContactEmail,
  LC.contactClassification
FROM
  Location_contacts AS LC
WHERE
  (
    (
      (LC.contactEmail) Is Not Null
    )
    AND (
      (LC.contactClassification) Like "TR*"
      Or (LC.contactClassification) Like "*Trust*"
    )
  )
ORDER BY
  LC.clientRef,
  LC.contactEmail;
