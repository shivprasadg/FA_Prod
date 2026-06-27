SELECT
  DISTINCT Clients.clientID,
  Clients.ClientShNm,
  Contacts.contactID,
  Contacts.contactClassification,
  Clients.clienttype,
  Contacts.[C-Level],
  Contacts.contactTitle,
  Notes.CallType,
  IIf(
    [contactclassification] = "signator",
    0.1, 0
  ) AS Sig,
  IIf(
    [c-level] = Yes
    And [calltype] = "phone call",
    0.1, 0
  ) AS CCall,
  IIf(
    [c-level] = Yes
    And [calltype] = "on-site",
    0.2, 0
  ) AS CVisit,
  IIf(
    [c-level] = No
    And [calltype] = "phone call",
    0.05, 0
  ) AS ACall,
  IIf(
    [c-level] = No
    And [calltype] = "on-site",
    0.1, 0
  ) AS AVisit
FROM
  Clients
  INNER JOIN (
    Notes
    INNER JOIN Contacts ON Notes.MainContact = Contacts.contactID
  ) ON Clients.clientID = Notes.clientRef
WHERE
  (
    (
      (Contacts.contactClassification)= "signator"
    )
    AND (
      (Clients.clienttype)= "target"
    )
  )
  OR (
    (
      (Clients.clienttype)= "target"
    )
    AND (
      (Contacts.[C-Level])= Yes
    )
    AND (
      (Contacts.contactTitle) Like "*cfo*"
    )
    AND (
      (Notes.CallType)= "on-site"
    )
  )
  OR (
    (
      (Clients.clienttype)= "target"
    )
    AND (
      (Contacts.contactTitle) Like "*Analyst*"
    )
    AND (
      (Notes.CallType)= "on-site"
    )
  )
  OR (
    (
      (Contacts.contactClassification)= "financial"
      Or (Contacts.contactClassification)= "primary"
    )
    AND (
      (Clients.clienttype)= "target"
    )
    AND (
      (Contacts.[C-Level])= Yes
    )
    AND (
      (Contacts.contactTitle) Like "*cfo*"
    )
    AND (
      (Notes.CallType)= "phone call"
    )
  )
  OR (
    (
      (Contacts.contactClassification)= "financial"
      Or (Contacts.contactClassification)= "primary"
    )
    AND (
      (Clients.clienttype)= "target"
    )
    AND (
      (Contacts.contactTitle) Like "*Analyst*"
    )
    AND (
      (Notes.CallType)= "phone call"
    )
  );
