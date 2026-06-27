SELECT
  Contacts.contactID,
  SplitString([contactFullName], ".", 0)& "." AS Salutation,
  [contactFirst] & " " & IIf(
    Nz([contactMiddle], "")= "",
    Null,
    [contactMiddle] & " "
  )& [ContactLast] AS FullName,
  Contacts.contactTitle AS Title,
  Contacts.contactPrimaryPhone AS Phone,
  Contacts.contactCell AS Mobile,
  IIf(
    Nz([contactEmail], "")<> "",
    "Y",
    "N"
  ) AS HasEmail,
  Contacts.contactEmail AS EmailAddr,
  Contacts.clientRef AS ClientID,
  Contacts.contactFullName
FROM
  Contacts;
