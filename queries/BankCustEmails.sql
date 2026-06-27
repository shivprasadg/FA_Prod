SELECT
  Min(BankCustEmails0.CustName) AS CustName,
  Min(BankCustEmails0.CustEmail) AS CustEmail,
  Min(BankCustEmails0.CustCName) AS CustCName,
  Min(BankCustEmails0.CustCEmail) AS CustCEmail,
  Min(BankCustEmails0.ABSEmail) AS ABSEmail
FROM
  (
    SELECT
      Location_contacts.clientRef AS BankId,
      IIf(
        [contactClassification] = "Custodian",
        [contactFullName], Null
      ) AS CustName,
      IIf(
        [contactClassification] = "Custodian",
        [contactEmail], Null
      ) AS CustEmail,
      IIf(
        [contactClassification] = "Custodian-Contact",
        [contactFullName], Null
      ) AS CustCName,
      IIf(
        [contactClassification] = "Custodian-Contact",
        [contactEmail], Null
      ) AS CustCEmail,
      IIf(
        [contactClassification] = "Custodian-Vault",
        [contactEmail], Null
      ) AS ABSEmail
    FROM
      Location_contacts
    WHERE
      (
        (
          (Location_contacts.clientRef)= 5233
        )
        AND (
          (
            Location_contacts.contactClassification
          ) Like "*Custodian*"
        )
      )
    ORDER BY
      Location_contacts.clientRef,
      IIf(
        [contactClassification] = "Custodian",
        [contactEmail], Null
      )
  ) AS BankCustEmails0;
