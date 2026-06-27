SELECT
  Location_contacts.clientRef AS BankId,
  IIf(
    [contactClassification] = "Primary Trust Incumbent",
    [contactEmail], Null
  ) AS PrimaryTrustIncumbentEmail,
  IIf(
    [contactClassification] = "TR Incumbent",
    [contactFullName], Null
  ) AS TrustIncumbentName,
  IIf(
    [contactClassification] = "TR Incumbent",
    [contactEmail], Null
  ) AS TrustIncumbentEmail,
  IIf(
    [contactClassification] = "TR Incumbent-Contact",
    [contactFullName], Null
  ) AS TRIncumbentContactName,
  IIf(
    [contactClassification] = "TR Incumbent-Contact",
    [contactEmail], Null
  ) AS TRIncumbentContactEmail,
  IIf(
    [contactClassification] = "TR Contact",
    [contactFullName], Null
  ) AS TrustContactName,
  IIf(
    [contactClassification] = "TR Contact",
    [contactEmail], Null
  ) AS TrustContactEmail
FROM
  Location_contacts
WHERE
  (
    (
      (
        IIf(
          [contactClassification] = "Primary Trust Incumbent",
          [contactEmail], Null
        )
      ) Is Not Null
    )
  )
  OR (
    (
      (
        IIf(
          [contactClassification] = "TR Incumbent",
          [contactEmail], Null
        )
      ) Is Not Null
    )
  )
  OR (
    (
      (
        IIf(
          [contactClassification] = "TR Incumbent-Contact",
          [contactEmail], Null
        )
      ) Is Not Null
    )
  )
  OR (
    (
      (
        IIf(
          [contactClassification] = "TR Contact",
          [contactEmail], Null
        )
      ) Is Not Null
    )
  )
ORDER BY
  IIf(
    [contactClassification] = "Primary Trust Incumbent",
    [contactEmail], Null
  ) DESC;
