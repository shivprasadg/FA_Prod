Select
  ClientID,
  MLID,
  SignerGroup1 As SignerType,
  SchID,
  SGrpID,
  MLSignerID As ContactID
FROM
  ATF_FAID_ContactsMLSigners
WHERE
  SignerGroup1 = 'Master Lease'
Union ALL
Select
  ClientID,
  MLID,
  SignerGroup1 As SignerType,
  SchID,
  SGrpID,
  MLSigner2ID As ContactID
FROM
  ATF_FAID_ContactsMLSigners
WHERE
  SignerGroup1 = 'Master Lease'
Union ALL
Select
  ClientID,
  MLID,
  SignerGroup1 As SignerType,
  SchID,
  SGrpID,
  MLSigner3ID As ContactID
FROM
  ATF_FAID_ContactsMLSigners
WHERE
  SignerGroup1 = 'Master Lease'
Union ALL
Select
  ClientID,
  MLID,
  SignerGroup2 As SignerType,
  SchID,
  SGrpID,
  MLGuarSignerID As ContactID
FROM
  ATF_FAID_ContactsMLSigners
WHERE
  SignerGroup2 = 'Guarantor';
Union ALL
Select
  ClientID,
  MLID,
  SignerGroup3 As SignerType,
  SchID,
  SGrpID,
  SchSignerID As ContactID
FROM
  ATF_FAID_ContactsMLSigners
WHERE
  SignerGroup3 = 'Schedule';
Union ALL
Select
  ClientID,
  MLID,
  SignerGroup3 As SignerType,
  SchID,
  SGrpID,
  SchSigner2ID As ContactID
FROM
  ATF_FAID_ContactsMLSigners
WHERE
  SignerGroup3 = 'Schedule';
Union ALL
Select
  ClientID,
  MLID,
  SignerGroup3 As SignerType,
  SchID,
  SGrpID,
  SchSigner3ID As ContactID
FROM
  ATF_FAID_ContactsMLSigners
WHERE
  SignerGroup3 = 'Schedule';
UNION ALL
Select
  ClientID,
  MLID,
  SignerGroup4 As SignerType,
  SchID,
  SGrpID,
  CofASignerID As ContactID
FROM
  ATF_FAID_ContactsMLSigners
WHERE
  SignerGroup4 = 'COA';
